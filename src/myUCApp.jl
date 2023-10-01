module myUCApp
using JuMP,DataFrames,CPLEX,XLSX,Dates,Suppressor,JSON,NativeFileDialog,UnicodePlots,PrettyTables
export UC
function UC()::Cint
    try
        config = JSON.parsefile("assets/config.json")
        OFFSET = 7 #从第二天的7点开始规划
        function myfun(r)
            v = r[:时间]
            return v == "高峰电量"
        end
        NOW = Int(ceil(hour(now())*4 + minute(now())/15))
        # NOW = 75
        # filename = "assets/日前计划编制报表20220721.xlsx"
        #获取当前超短期数据
        today_schedule = DataFrame(XLSX.readtable("assets/today_schedule.xlsx",1,"B:K";
        first_row=2)...)[2:end,:]
        # println("请选择日前计划文件（仅支持xlsx格式）！")
        # filename = pick_file("";filterlist="xlsx")
        filename = "assets/next_day_schedule.xlsx"
        next_day_schedule = DataFrame(XLSX.readtable(filename,1,"A:EG";
        first_row=3,stop_in_row_function=myfun)...)
        next_day_string = replace(split(XLSX.readdata(filename,1,"AJ2"))[1],r"[年月日]"=>"")
        next_day = DateTime(next_day_string,DateFormat("yyyymmdd"))
        row_string = XLSX.readtable(filename, 1,"A:EG";
        first_row=3)[1][2][end]
        re = r"燃机开机容量.*万方"
        gas_plan_string = collect(eachmatch(re,row_string))[1].match
        total_capacity = match(r"燃机开机容量(?P<total>\d*)万千瓦",gas_plan_string)["total"]
        # gas_plan = collect(eachmatch(r"[\u4e00-\u9fa5]*[1-9]{1}[机组]",gas_plan_string))
        gas_plan = collect(eachmatch(r"(?P<plant>[\u4e00-\u9fa5]*)(?P<plan>[1-9]{1}[机组][^，]*)[（\(](?P<gas>\d*)万方",gas_plan_string))
        # gas_plan_tmp = [[c for c in x.match] for x in gas_plan]
        # gas_plan_unpack = Array([(join(x[1:end-2],""),parse(Int,x[end-1])) for x in gas_plan_tmp])
        namemap = DataFrame(XLSX.readtable("assets/燃机名称表.xlsx","别名","A:B")...)
        parametersmap = DataFrame(XLSX.readtable("assets/燃机名称表.xlsx","参数","A:F")...)
        ST_COST = DataFrame(XLSX.readtable("assets/燃机名称表.xlsx","启动成本","A:F")...)[!,:启动成本]
        # gas_total = []
        # for s in eachmatch(r"[\(（][1-9]\d*万方",gas_plan_string)
        #     push!(gas_total,parse(Int,match(r"[1-9]\d*",s.match).match))
        # end
        alias_to_full = Dict()
        parameters = Dict()
        for i in range(1,stop=nrow(namemap))
            alias_to_full[namemap[i,:别名]] = namemap[i,:电厂]
        end
        for i in range(1,stop=nrow(parametersmap))
            parameters[parametersmap[i,:电厂]] = Dict()
            parameters[parametersmap[i,:电厂]]["Pmax"] = parametersmap[i,:单机出力上限]
            parameters[parametersmap[i,:电厂]]["Pmin"] = parametersmap[i,:单机出力下限]
            parameters[parametersmap[i,:电厂]]["unit_gas"] = parametersmap[i,:度电耗气]
            parameters[parametersmap[i,:电厂]]["输出名称"] = parametersmap[i,:输出名称]
        end
        param = Dict()
        plant = Dict()
        ordered_unit = []
        # 获取燃气电厂当前气量
        scada_today = XLSX.readtable("assets/燃机名称表.xlsx","当日量测")
        for i in 2:length(scada_today[2])
            k = String(scada_today[2][i])#解析出来是Symbol，转化成为String
            remained_gas,current_power = scada_today[1][i][4],scada_today[1][i][6]
            if round(current_power/parameters[k]["Pmax"]) >= 1
                plant[k] = Dict()
                plant[k]["total_gas"] = 0
                plant[k]["remained_gas"] = remained_gas
                # 辨识当前燃机开机台数（下一步用可用容量试试）
                plant[k]["running_units"] = Int(round(current_power/parameters[k]["Pmax"]))
                plant[k]["units"] = []
            end
        end
        #解析明日气量和开机方式
        for this_plant in gas_plan
            k = alias_to_full[this_plant["plant"]]
            if k ∉ keys(plant) 
                plant[k] = Dict("total_gas"=>parse(Int,this_plant["gas"]),"units"=>[],"running_units"=>0,"remained_gas"=>0)
            else
                plant[k]["total_gas"] = parse(Int,this_plant["gas"])
            end
            if length(this_plant["plan"]) == 2 #无特殊要求，默认不过夜(这个分支应该调用不到，后续考虑删除)
                N = parse(Int,this_plant["plan"][1])
                for n in 1:N
                    param["$k#$(n)机"] = Dict()
                    param["$k#$(n)机"]["Pmax"] = parameters[k]["Pmax"]
                    param["$k#$(n)机"]["Pmin"] = parameters[k]["Pmin"]
                    param["$k#$(n)机"]["单耗"] = parameters[k]["unit_gas"]
                    param["$k#$(n)机"]["是否过夜"] = 0
                    param["$k#$(n)机"]["输出名称"] = parameters[k]["输出名称"] * "#$(n)机"
                    push!(plant[k]["units"],"$k#$(n)机")
                end
            else
                n = 1
                for mc in eachmatch(r"(?P<number>\d{1})[机组](?P<method>[\u4e00-\u9fa5]*)",this_plant["plan"])
                    for i in 1:parse(Int,mc["number"])
                        param["$k#$(n)机"] = Dict()
                        param["$k#$(n)机"]["Pmax"] = parameters[k]["Pmax"]
                        param["$k#$(n)机"]["Pmin"] = parameters[k]["Pmin"]
                        param["$k#$(n)机"]["单耗"] = parameters[k]["unit_gas"]
                        param["$k#$(n)机"]["输出名称"] = parameters[k]["输出名称"] * "#$(n)机"
                        param["$k#$(n)机"]["是否过夜"] = 0
                        param["$k#$(n)机"]["次日开机"] = true
                        push!(plant[k]["units"],"$k#$(n)机")
                        push!(ordered_unit,"$k#$(n)机")               
                        if n <= plant[k]["running_units"]
                            param["$k#$(n)机"]["running"] = 1
                            if mc["method"]=="过夜" || mc["method"]=="连续运行"
                                param["$k#$(n)机"]["是否过夜"] = 1
                            end
                        else
                            param["$k#$(n)机"]["running"] = 0
                        end
                        n += 1
                    end
                end
                #若n-1小于当前组数，说明有若干机组明天不开
                if n-1 < plant[k]["running_units"]
                    for i in n:plant[k]["running_units"]
                        param["$k#$(i)机"] = Dict()
                        param["$k#$(i)机"]["Pmax"] = parameters[k]["Pmax"]
                        param["$k#$(i)机"]["Pmin"] = parameters[k]["Pmin"]
                        param["$k#$(i)机"]["单耗"] = parameters[k]["unit_gas"]
                        param["$k#$(i)机"]["输出名称"] = parameters[k]["输出名称"] * "#$(i)机"
                        param["$k#$(i)机"]["running"] = 1
                        param["$k#$(i)机"]["是否过夜"] = 0
                        param["$k#$(i)机"]["次日开机"] = false
                        push!(plant[k]["units"],"$k#$(i)机")
                        push!(ordered_unit,"$k#$(i)机")
                    end
                end
            end    
        end
        for k in keys(plant)
            if length(plant[k]["units"]) == 0
                for i in 1:plant[k]["running_units"]
                    param["$k#$(i)机"] = Dict()
                    param["$k#$(i)机"]["Pmax"] = parameters[k]["Pmax"]
                    param["$k#$(i)机"]["Pmin"] = parameters[k]["Pmin"]
                    param["$k#$(i)机"]["单耗"] = parameters[k]["unit_gas"]
                    param["$k#$(i)机"]["输出名称"] = parameters[k]["输出名称"] * "#$(i)机"
                    param["$k#$(i)机"]["running"] = 1
                    param["$k#$(i)机"]["是否过夜"] = 0
                    param["$k#$(i)机"]["次日开机"] = false
                    push!(plant[k]["units"],"$k#$(i)机")
                    push!(ordered_unit,"$k#$(i)机")
                end
            end 
        end
        T = 96 + 32
        T0 = NOW - 96
        df = DataFrame()
        df[!,"时刻"] = [next_day + Minute(t*15) for t in T0:T]
        for x in ordered_unit
            if param[x]["是否过夜"] == 1
                df[!,x] = [t ∈ 1:32 ? 1 : missing for t in T0:T]
            else
                df[!,x] = [missing for t in T0:T]
            end
        end
        XLSX.writetable("assets/必开燃机列表.xlsx",df;overwrite=true)
        println("请设置必开燃机，设置完后保存关闭，按回车键继续")
        if Sys.iswindows()
            run(`cmd /c .\\assets\\必开燃机列表.xlsx`)
        else
            run(`open ./assets/必开燃机列表.xlsx`)
        end
        # readline()
        required_uc = DataFrame(XLSX.readtable("assets/必开燃机列表.xlsx",1)...)
        # println([x for x in keys(param) if param[x]["running"]])
        for p in keys(plant)
            println("$(p)当前剩余气量$(round(plant[p]["remained_gas"];digits=1))万方")
        end
        # ## 必开燃机中过夜
        # for x in keys(param)
        #     if !ismissing(sum(required_uc[!,x]))
        #         if sum(required_uc[!,x]) == 96
        #             param[x]["是否过夜"] = 1
        #         end
        #     end
        # end
        nuclear,wind,solar,ext,load = next_day_schedule[!,:核电],next_day_schedule[!,:风电],next_day_schedule[!,:光伏],next_day_schedule[!,:受电],next_day_schedule[!,:统调负荷]
        gas_power = next_day_schedule[!,:燃气]
        nuclear0,wind0,solar0,ext0,load0 = 
        parse.(Float64,today_schedule[!,:核电]),
        parse.(Float64,today_schedule[!,:风电]),
        parse.(Float64,today_schedule[!,:光伏]),
        parse.(Float64,today_schedule[!,:受电]),
        - parse.(Float64,today_schedule[!,:超短期负荷])
        # 拼接次日24:00-后天7:00的计划类数据
        for list in (nuclear,wind,solar,ext,load,gas_power)
            for i in 1:32
                push!(list,list[i])
            end
        end
        COAL_PMAX,COAL_PMIN = config["统调燃煤机组最大发电能力"],config["统调燃煤机组最小发电能力"]
        RESERVE0,RESERVE1 = config["应留备用"],config["期望备用"]
        WATER = config["水电发电能力"]
        PR = config["成本系数"]
        m = Model(CPLEX.Optimizer)
        println("正在进行备用优化...") 
        ### 
        @variable(m,st[x in keys(param),t in T0:T],binary=true)
        @variable(m,up[x in keys(param),t in T0:T],binary=true)
        @variable(m,down[x in keys(param),t in T0:T],binary=true)
        @variable(m,hot[x in keys(param),t in T0:T],binary=true)
        @variable(m,cold[x in keys(param),t in T0:T],binary=true)
        # @variable(m,UT[x in keys(param),t in 0:T])
        # @variable(m,DT[x in keys(param),t in 0:T])
        @variable(m,Pg[x in keys(param),t in T0:T],lower_bound=0)
        @variable(m,dPg[x in keys(param),t in T0:T],lower_bound=0)
        @variable(m,PMg[x in keys(param),t in T0:T],lower_bound=0)
        @variable(m,lackreserve[t in T0:T],lower_bound=0)#正备用缺口惩罚
        @variable(m,windcut[t in T0:T],lower_bound=0)#负备用缺口
        @variable(m,rho)
        @variable(m,reserve[t in T0:T])
        @variable(m,neg_reserve[t in T0:T])
        @variable(m,dg1_plus[p in keys(plant)],lower_bound=0)
        @variable(m,dg1_minus[p in keys(plant)],lower_bound=0)
        @variable(m,dg1_indicator[p in keys(plant)],binary=true) #进行挪气的指示变量
        @variable(m,dg2[p in keys(plant)],lower_bound=0)
        @variable(m,dg_total_plus,lower_bound=0)
        @variable(m,dg_total_minus,lower_bound=0)
        # @variable(m,coal[t in 1:T],lower_bound=COAL_PMIN,upper_bound=COAL_PMAX)
        # @variable(m,water[t in 1:T],lower_bound=0,upper_bound=WATER)
        # @variable(m,fuel_cost[t in 1:T],lower_bound=0)
        gas_cons_ref = []
        for t in T0:T
            ### 机组约束
            for x in keys(param)
                ## 必开必停
                if !ismissing(required_uc[t-T0+1,x])
                    @constraint(m,st[x,t] == required_uc[t-T0+1,x])
                end
                if !PR["允许凌晨起机"]
                    if ST_COST[mod(t,96)+1] >= 10000
                        @constraint(m,up[x,t] == 0)
                    end
                end
                ## 次日启动机组最大连续停机时间约束
                if t <= T - 64 && t >= 32 && param[x]["次日开机"]
                    @constraint(m,sum(st[x,s] for s in t:t+52) >= 1)
                end
                ## 禁止停机
                if t <= T0 + config["禁止时段"]
                    @constraint(m,down[x,t] == 0)
                end
                ## 机组技术约束
                @constraint(m,Pg[x,t] <= (1-cold[x,t])*param[x]["Pmax"])# 技术出力上限
                @constraint(m,Pg[x,t] >= hot[x,t]*param[x]["Pmin"])# 技术出力下限
                @constraint(m,PMg[x,t] <= Pg[x,t] + hot[x,t] * param[x]["Pmax"])
                @constraint(m,PMg[x,t] >= Pg[x,t] - hot[x,t] * param[x]["Pmax"])
                @constraint(m,PMg[x,t] <= param[x]["Pmax"])
                @constraint(m,PMg[x,t] >= hot[x,t] * param[x]["Pmax"])
                # @constraint(m,Pg[x,t] <= st[x,t]*param[x]["Pmax"])# 技术出力上限
                # @constraint(m,Pg[x,t] >= st[x,t]*param[x]["Pmin"])# 技术出力下限
    
                # ##启动曲线
                # @constraint(m,Pg[x,t] >= UT[x,t]/4 * param[x]["Pmin"] - hot[x,t]*10000 - (1-st[x,t])*10000)
                # @constraint(m,Pg[x,t] <= UT[x,t]/4 * param[x]["Pmin"] + hot[x,t]*10000 + (1-st[x,t])*10000)
                # ##
                # ##停机曲线
                # @constraint(m,Pg[x,t] >= (4-DT[x,t])/4 * param[x]["Pmin"] - cold[x,t]*10000 - st[x,t]*10000)
                # @constraint(m,Pg[x,t] <= (4-DT[x,t])/4 * param[x]["Pmin"] + cold[x,t]*10000 + st[x,t]*10000)
                # #
                if t >= T0+1 #爬坡率约束
                    @constraint(m,Pg[x,t] - Pg[x,t-1] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                    @constraint(m,Pg[x,t-1] - Pg[x,t] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                # else
                #     if param[x]["running"] == 0
                #         @constraint(m,Pg[x,t] <= hot[x,t]*param[x]["Pmax"]/3 + (1-hot[x,t])*param[x]["Pmin"]/4)
                #     end
                end
                if t >= T0+1 #爬坡成本约束
                    @constraint(m,dPg[x,t] >= Pg[x,t] - Pg[x,t-1])
                    @constraint(m,dPg[x,t] >= Pg[x,t-1] - Pg[x,t])
                end
                if t >= T0+1
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - st[x,t-1])
                else
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - param[x]["running"])
                end
                if param[x]["running"] == 1 && t <= 12
                    @constraint(m,up[x,t] == 0)
                end
                if param[x]["running"] == 0 && t <= 0 
                    @constraint(m,up[x,t] == 0)
                    @constraint(m,down[x,t] == 0)
                end
                @constraint(m,up[x,t] + down[x,t] <= 1)
                @constraint(m,up[x,t] <= 1 - sum(down[x,t-s] for s in 0:min(t-T0,20)))#停机后间隔四个小时才能开机
                @constraint(m,down[x,t] <= 1 - sum(up[x,t-s] for s in 0:min(t-T0,20)))#开机后间隔四个小时才能停机
                @constraint(m, hot[x,t] == st[x,t] - sum(up[x,t-s] for s in 0:min(t-T0,3)))#启动好了
                @constraint(m, cold[x,t] == 1 - st[x,t] - sum(down[x,t-s] for s in 0:min(t-T0,3)))#停好了
            end
            ### 机组约束
            # ## 煤机总体约束
            # for k in 0:20
            #     @constraint(m,fuel_cost[t] >= 0.25 * 100* (1 + k/20) * (coal[t] - (k/20 * COAL_PMAX + (20-k)/20 * COAL_PMIN)))
            # end
            ## 电网约束
            if t >= 1 #次日和后天平衡
                # @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
                # sum(PMg[x,t] for x in keys(param)) >= load[t] + RESERVE - loadcut[t]) #正备用约束
                @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
                sum(PMg[x,t] for x in keys(param)) == load[t] + reserve[t]) #正备用约束
                @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
                sum(Pg[x,t] for x in keys(param)) <= load[t] + windcut[t]) #负备用约束
                @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
                sum(Pg[x,t] for x in keys(param)) == load[t] - neg_reserve[t]) #负备用约束
                @constraint(m,rho <= reserve[t])
            else
                # @constraint(m,COAL_PMAX + WATER + sum(l[96+t] for l in [nuclear0,wind0,solar0,ext0]) +
                # sum(PMg[x,t] for x in keys(param)) >= load0[96+t] + RESERVE - loadcut[t]) #正备用约束
                @constraint(m,COAL_PMAX + WATER + sum(l[96+t] for l in [nuclear0,wind0,solar0,ext0]) +
                sum(PMg[x,t] for x in keys(param)) == load0[96+t] + reserve[t]) #正备用约束
                @constraint(m,COAL_PMIN + sum(l[96+t] for l in [nuclear0,wind0,solar0,ext0]) +
                sum(Pg[x,t] for x in keys(param)) <= load0[96+t] + windcut[t]) #负备用约束
                @constraint(m,COAL_PMIN + sum(l[96+t] for l in [nuclear0,wind0,solar0,ext0]) +
                sum(Pg[x,t] for x in keys(param)) == load0[96+t] - neg_reserve[t]) #负备用约束
                @constraint(m,rho <= reserve[t])
            end
            @constraint(m,lackreserve[t] >= PR["小于应留备用惩罚"]*(RESERVE0 - reserve[t]))
            @constraint(m,lackreserve[t] >= PR["小于期望备用惩罚"]*(RESERVE1 - reserve[t]))
        end
        # for x in keys(param)# TODO: 启停机的额外要求
        #     # if param[x]["是否过夜"] == 1
        #     #     @constraint(m,sum(up[x,t] for t in 1:T) == 0)
        #     #     @constraint(m,sum(down[x,t] for t in 1:T) == 0)
        #     # else
        #     @constraint(m,sum(up[x,t] for t in 1:T) <= config["最多启停次数"]) #日间限制启动次数
        # end
        for x in keys(param)
            @constraint(m,sum(up[x,t] for t in 1:96) <= config["最多启停次数"]) #日间限制启动次数
        end
        for p in keys(plant) #总气量约束
            @constraint(m,0.1*sum(Pg[x,t]*0.25*param[x]["单耗"] for x in plant[p]["units"] for t in T0:32) + dg1_plus[p] - dg1_minus[p] == plant[p]["remained_gas"])
            @constraint(m,0.1*sum(Pg[x,t]*0.25*param[x]["单耗"] for x in plant[p]["units"] for t in 33:T) + dg2[p] == plant[p]["total_gas"])
        end
        #挪气平衡约束
        if config["凌晨挪气"]
            @constraint(m,sum(dg1_plus) - sum(dg1_minus) + dg_total_plus - dg_total_minus == 0)
            for p in keys(plant)
                @constraint(m,dg1_indicator[p] >= (dg1_plus[p] + dg1_minus[p])/1000)
            end
            @objective(m,Min,0.25*sum(lackreserve[t] for t in T0:96) + 
                        PR["后一日凌晨系数"]*0.25*sum(lackreserve[t] for t in 97:T) + 
                        # 0.1*sum(dPg) + 
                        PR["当日气量偏差惩罚"] * (dg_total_minus + dg_total_plus) + 
                        PR["单机挪气惩罚"] * sum(dg1_indicator) + 
                        PR["固定成本"]*sum(st) + 
                        sum(up[x,t]*ST_COST[mod(t,96)+1] for t in T0:T for x in keys(param)) + 
                        sum(windcut) + 
                        PR["次日气量富余惩罚"]*sum(dg2))
        else
            @objective(m,Min,0.25*sum(lackreserve[t] for t in T0:96) + 
            PR["后一日凌晨系数"]*0.25*sum(lackreserve[t] for t in 97:T) + 
            # 0.1*sum(dPg) + 
            PR["固定成本"]*sum(st) + 
            sum(up[x,t]*ST_COST[mod(t,96)+1] for t in T0:T for x in keys(param)) + 
            sum(windcut) + 
            PR["当日气量偏差惩罚"]*sum(dg1_plus) + 
            PR["当日气量偏差惩罚"]*sum(dg1_minus) + 
            PR["次日气量富余惩罚"]*sum(dg2))
        end
        # @objective(m,Min,10000*0.25*sum(loadcut[t] for t in T0:96) + 
        #             5000*0.25*sum(loadcut[t] for t in 97:T) + 
        #             sum(windcut) + 
        #             10000*sum(dg1) + 
        #             10000*sum(dg2))
        set_time_limit_sec(m, config["maxTimeforUC"])
        if config["logUC"]
            optimize!(m)
        else
            @suppress_out optimize!(m)
        end
        df = DataFrame()
        df[!,"时刻"] = [next_day + Minute(t*15) for t in T0:T]
        df[!,"煤机最大发电能力"] = [COAL_PMAX for t in T0:T]
        df[!,"煤机最小发电能力"] = [COAL_PMIN for t in T0:T]
        df[!,"燃机最大发电能力"] = Int.(round.([sum(value(PMg[x,t]) for x in keys(param)) for t in T0:T]))
        df[!,"水电"] = [WATER for t in T0:T]
        df[!,"核电"] = [t >= 1 ? nuclear[t] : nuclear0[96+t] for t in T0:T]
        df[!,"风电"] = [t >= 1 ? wind[t] : wind0[96+t] for t in T0:T]
        df[!,"光伏"] = [t >= 1 ? solar[t] : solar0[96+t] for t in T0:T]
        df[!,"受电"] = [t >= 1 ? ext[t] : ext0[96+t] for t in T0:T]
        df[!,"统调负荷"] =[t >= 1 ? load[t] : load0[96+t] for t in T0:T]
        df[!,"正备用"] = Int.(round.([value(reserve[t]) for t in T0:T]))
        df[!,"负备用"] = Int.(round.([value(neg_reserve[t]) for t in T0:T]))
        df[!,"燃机总出力"] = Int.(round.([sum(value(Pg[x,t]) for x in keys(param)) for t in T0:T]))
        for x in ordered_unit
            df[!,x] = Int.(round.([value(Pg[x,t]) for t in T0:T]))
        end
        df2 = DataFrame()
        df2[!,"燃气电厂"] = [parameters[p]["输出名称"] for p in keys(plant)]
        df2[!,"计划气量"] = [plant[p]["total_gas"] for p in keys(plant)]
        df3 = DataFrame()
        df3[!,"时刻"] = [Dates.format(next_day + Minute(t*15),dateformat"yyyy-mm-dd HH:MM:SS") for t in 1:96]
        for x in ordered_unit
            df3[!,param[x]["输出名称"]] = Int.(round.([value(Pg[x,t]) for t in 1:96]))
        end
        println("优化成功！")
        println("正备用最低值为$(minimum(value.(reserve)))")
        println("================================================")
        println("当日/次日凌晨燃机计划：")
        for p in keys(plant)
            if sum(param[x]["running"] for x in plant[p]["units"]) == 0
                continue
            end
            print("$p:  ")
            if value(dg1_plus[p]) + value(dg1_minus[p]) >= 0.6
                if value(dg1_plus[p]) > 0
                    print("（减少$(Int(round(value(dg1_plus[p]))))万方）")
                else 
                    print("（增加$(Int(round(value(dg1_minus[p]))))万方）")
                end
            end
            println()
            for x in plant[p]["units"]
                if param[x]["running"] == 1
                    print("    $(x[end-2:end]) ")
                    if sum(value(down[x,t]) for t in T0:32) >= 1
                        for t in T0:32
                            if round(value(down[x,t])) == 1
                                dt = Dates.format(next_day + Minute(t*15),"HH:MM")
                                if t <= 0
                                    print("当日$(dt)停机")
                                else
                                    print("凌晨$(dt)停机")
                                end
                            end
                        end
                    else
                        print("凌晨过夜运行")
                    end
                    println()
                end
            end
        end
        println("================================================")
        println("次日燃机开机计划")
        for p in keys(plant)
            if sum(value(up[x,t]) for x in plant[p]["units"] for t in 1:96) >= 1
                println("$p ：$(plant[p]["total_gas"])万方")
                for x in plant[p]["units"]
                    print("     $(x[end-2:end])  ")
                    for t in 1:T
                        if round(value(up[x,t])) == 1
                            if t <= 96
                                dt = Dates.format(next_day + Minute(t*15) + Hour(1),"HH:MM")
                                print("$(dt)带足----") 
                            end
                        end
                        if round(value(down[x,t])) == 1
                            if t >= 33
                                dt = Dates.format(next_day + Minute(t*15),"HH:MM")
                                if t<=95
                                    print("$(dt)停机    ")
                                else
                                    print("后一日$(dt)停机    ")
                                end
                            end
                        end
                    end
                    print("\n")
                end
            end
        end

        println("================================================")
        println("当前至次日24点旋备曲线")
        plt = lineplot(T0:T, df[!,"正备用"], name="optimized",
            xlabel="time", ylabel="spinning reserve(MW)", canvas=DotCanvas, border=:ascii,xlim=(T0,T))
        lineplot!(plt, T0:T, [RESERVE0 for t in T0:T], name="required",
        color=:cyan)
        display(plt)
        println("================================================")
        println("当前至次日24点燃机备用情况")
        plt2 = lineplot(T0:T,df[!,"燃机最大发电能力"], name="optimized",
        xlabel="time", ylabel="spinning reserve(MW)", canvas=DotCanvas, border=:ascii,xlim=(T0,T))
        lineplot!(plt2,T0:T,max.((df[!,"燃机最大发电能力"] - df[!,"正备用"] .+ RESERVE0),0), name="required",
        color=:cyan)
        display(plt2)
        # plants_keys = collect(keys(plant))
        # for i in 1:length(plants_keys)
        #     println("$(plants_keys[i]) 的气量影子价格为:$(dual(gas_cons_ref[i]))")
        # end
        # println("明日备用情况：按（手填）曲线考虑")
        df_new = df[2-T0:97-T0,names(df)]
        data = []
        println("================================================")
        println("次日正备用最紧张时段分析")
        for (timeslot,t1,t2) in zip(["凌晨","早峰","午峰","晚峰"],[1,28,44,72],[27,43,71,96])
            tmpdf = df_new[t1:t2,names(df)]
            lowest,p = findmin(tmpdf[!,"正备用"])
            tl,tc,tr,te,tso,twi = round(tmpdf[!,"统调负荷"][p]/10),round(tmpdf[!,"煤机最大发电能力"][p]/10),round(tmpdf[!,"燃机总出力"][p]/10),round(tmpdf[!,"受电"][p]/10),round(tmpdf[!,"光伏"][p]/10),round(tmpdf[!,"风电"][p]/10)
            tmg = round(tmpdf[!,"燃机最大发电能力"][p]/10)
            lowest_time = Dates.format(tmpdf[!,"时刻"][p],"HH:MM")
            push!(data,permutedims([timeslot,lowest_time,lowest,Int(tl),Int(tc),Int(tmg),Int(te),Int(tso),Int(twi)]))
            # println("$(timeslot)，$(lowest_time)备用最紧，为$(Int(round(lowest/10)))万千瓦；该时刻，统调负荷$(Int(tl))万千瓦，煤机$(Int(tc))万千瓦，燃机$(Int(tmg))万千瓦，受电$(Int(te))万千瓦，统调光伏$(Int(tso))万千瓦，统调风电$(Int(twi))万千瓦。")
        end
        data = vcat(data...)
        pretty_table(data;header=(["时段","时刻","正备用","统调负荷","燃煤","燃气","受电","光伏","风电"],["","","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)"]),header_crayon=crayon"yellow bold",tf=tf_unicode_rounded)
        println("================================================")
        data = []
        println("次日负备用最紧张时段分析")
        for (timeslot,t1,t2) in zip(["凌晨","早峰","午峰","晚峰"],[1,28,44,72],[27,43,71,96])
            tmpdf = df_new[t1:t2,names(df)]
            lowest,p = findmin(tmpdf[!,"负备用"])
            tl,tc,tr,te,tso,twi = round(tmpdf[!,"统调负荷"][p]/10),round(tmpdf[!,"煤机最小发电能力"][p]/10),round(tmpdf[!,"燃机总出力"][p]/10),round(tmpdf[!,"受电"][p]/10),round(tmpdf[!,"光伏"][p]/10),round(tmpdf[!,"风电"][p]/10)
            lowest_time = Dates.format(tmpdf[!,"时刻"][p],"HH:MM")
            push!(data,permutedims([timeslot,lowest_time,lowest,Int(tl),Int(tc),Int(tr),Int(te),Int(tso),Int(twi)]))
            # println("$(timeslot)，$(lowest_time)备用最紧，为$(Int(round(lowest/10)))万千瓦；该时刻，统调负荷$(Int(tl))万千瓦，煤机$(Int(tc))万千瓦，燃机$(Int(tmg))万千瓦，受电$(Int(te))万千瓦，统调光伏$(Int(tso))万千瓦，统调风电$(Int(twi))万千瓦。")
        end
        data = vcat(data...)
        pretty_table(data;header=(["时段","时刻","负备用","统调负荷","燃煤","燃气","受电","光伏","风电"],["","","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)","(万kW)"]),header_crayon=crayon"yellow bold",tf=tf_unicode_rounded)
        # println("明日燃机开机容量$(total_capacity)万千瓦，总气量$(Int(sum(df2[!,"计划气量"])))万方。")
        # for this_plant in gas_plan
        #     k = this_plant["plant"]
        #     tmpstr = ""
        #     gy,ngy = sum(param[x]["是否过夜"] == 1 for x in plant[k]["units"]),sum(param[x]["是否过夜"] == 0 for x in plant[k]["units"])
        #     if ngy > 0 
        #         tmpstr *= "$(ngy)机日开夜停"
        #     end
        #     if gy > 0
        #         tmpstr *= "$(gy)机过夜"
        #     end
        #     print("$k $(tmpstr)（$(plant[k]["total_gas"])万方）。")
        #     Tplus7 = [(t + 26)%T + 1 for t in 1:T]
        #     if length(plant[k]["units"]) == 1
        #         x = plant[k]["units"][1]
        #         if param[x]["是否过夜"] == 1
        #             load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
        #             if load_rate <= 90
        #                 print("（需压出力运行）\n")
        #             else
        #                 print("\n")
        #             end
        #         else
        #             start,stop = findfirst([value(hot[x,t])>0 for t in 1:T]),findlast(([value(hot[x,t])>0 for t in 1:T]))
        #             if isnothing(start)
        #                 continue
        #             end
        #             start,stop = df[!,"时刻"][Tplus7[start]],df[!,"时刻"][Tplus7[stop]]
        #             load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
        #             if hour(stop) > 8
        #                 print("$(Dates.format(start,"HH:MM"))带足--$(Dates.format(stop,"HH:MM"))")
        #             else
        #                 print("$(Dates.format(start,"HH:MM"))带足--次日$(Dates.format(stop,"HH:MM"))")
        #             end
        #             if load_rate <= 95
        #                 print("（需压出力运行）\n")
        #             else
        #                 print("\n")
        #             end
        #         end
        #     else
        #         print("\n")
        #         for i in 1:length(plant[k]["units"])
        #             x = plant[k]["units"][i]
        #             if param[x]["是否过夜"] == 1
        #                 load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
        #                 print("     #$(i)机过夜")
        #                 if load_rate <= 90
        #                     print("（需压出力运行）\n")
        #                 else
        #                     print("\n")
        #                 end
        #             else
        #                 start,stop = findfirst([value(hot[x,t])>0 for t in 1:T]),findlast(([value(hot[x,t])>0 for t in 1:T]))
        #                 start,stop = df[!,"时刻"][Tplus7[start]],df[!,"时刻"][Tplus7[stop]]
        #                 load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
        #                 if hour(stop) > 8
        #                     print("     #$(i)机: $(Dates.format(start,"HH:MM"))带足--$(Dates.format(stop,"HH:MM"))")
        #                 else
        #                     print("     #$(i)机: $(Dates.format(start,"HH:MM"))带足--次日$(Dates.format(stop,"HH:MM"))")
        #                 end
        #                 if load_rate <= 90
        #                     print("（需压出力运行）\n")
        #                 else
        #                     print("\n")
        #                 end
        #             end
        #         end
        #     end
        # end
        println("请将燃机计划保存至本地，保存后手动另存为xls格式，可在停电智能管控平台-调度日计划（手动导入）中上传、通知...")
        save_file_name = save_file("";filterlist="xlsx")
        # save_file_name = save_dialog("结果保存至...", GtkNullContainer(), (GtkFileFilter("*.xlsx", name="All supported formats"), "*.xlsx"))
        # XLSX.writetable(save_file_name,df;overwrite=true)
        XLSX.writetable(save_file_name,"REPORT_A"=>df,"REPORT_B"=>df2,"REPORT_C"=>permutedims(df3,"时刻");overwrite=true)
        XLSX.openxlsx("assets/次日平衡模板.xlsx",mode="rw") do xf
            sheet = xf[1]
            for k in 1:96
                sheet["B$(k+1)"] = df[k-T0+1,"煤机最大发电能力"]
                sheet["D$(k+1)"] = df[k-T0+1,"水电"]
                sheet["E$(k+1)"] = 80
                sheet["F$(k+1)"] = df[k-T0+1,"核电"]
                sheet["G$(k+1)"] = df[k-T0+1,"风电"]
                sheet["G$(k+1)"] = df[k-T0+1,"风电"]
                sheet["H$(k+1)"] = df[k-T0+1,"光伏"]
                sheet["I$(k+1)"] = df[k-T0+1,"受电"]
                sheet["J$(k+1)"] = 0
                sheet["M$(k+1)"] = 0
                sheet["L$(k+1)"] = df[k-T0+1,"统调负荷"]
                sheet["P$(k+1)"] = df[k-T0+1,"煤机最小发电能力"]
            end
            for i in 20:50
                for j in 1:97
                    sheet[j,i] = missing
                end
            end
            for (j,x) in enumerate(ordered_unit)
                sheet[1,j+19] = x
                for i in 1:96
                    sheet[i+1,j+19] = Int(round(value(Pg[x,i])))
                end
            end
        end
        println("流程结束，按回车键退出...")
        s = readline()
        println(s)
        return 0# if things finished successfully
    catch err
        showerror(stdout, err, catch_backtrace())
        println("流程结束，按回车键退出...")
        s = readline()
        print(s)
        return 0
    end
end
end # module  