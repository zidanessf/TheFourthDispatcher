module myUCApp
using JuMP,DataFrames,CPLEX,XLSX,Dates,Suppressor,JSON,NativeFileDialog
export maxReserve,UC
function maxReserve()::Cint
    try
        config = JSON.parsefile("assets/config.json")
        OFFSET = 7 #从第二天的7点开始规划
        function myfun(r)
            v = r[:时间]
            return v == "高峰电量"
        end
        # filename = "assets/日前计划编制报表20220714.xlsx"
        println("请选择日前计划文件（仅支持xlsx格式）！")
        filename = pick_file(""; filterlist="xlsx")
        next_day_schedule = DataFrame(XLSX.readtable(filename,1,"A:EG";
        first_row=3,stop_in_row_function=myfun)...)
        next_day_string = replace(split(XLSX.readdata(filename,1,"AJ2"))[1],r"[年月日]"=>"")
        next_day = DateTime(next_day_string,DateFormat("yyyymmdd"))
        row_string = XLSX.readtable(filename, 1,"A:EG";
        first_row=3)[1][2][end]
        re = r"燃机开机容量.*万方"
        gas_plan_string = collect(eachmatch(re,row_string))[1].match
        # gas_plan = collect(eachmatch(r"[\u4e00-\u9fa5]*[1-9]{1}[机组]",gas_plan_string))
        gas_plan = collect(eachmatch(r"(?P<plant>[\u4e00-\u9fa5]*)(?P<plan>[1-9]{1}[机组][^，]*)（(?P<gas>\d*)万方",gas_plan_string))
        # gas_plan_tmp = [[c for c in x.match] for x in gas_plan]
        # gas_plan_unpack = Array([(join(x[1:end-2],""),parse(Int,x[end-1])) for x in gas_plan_tmp])
        namemap = DataFrame(XLSX.readtable("assets/燃机名称表.xlsx","Sheet1","A:E")...)
        # gas_total = []
        # for s in eachmatch(r"[\(（][1-9]\d*万方",gas_plan_string)
        #     push!(gas_total,parse(Int,match(r"[1-9]\d*",s.match).match))
        # end
        alias_to_full = Dict()
        for i in range(1,stop=nrow(namemap))
            alias_to_full[namemap[i,:别名]] = Dict()
            alias_to_full[namemap[i,:别名]]["Pmax"] = namemap[i,:单机出力上限]
            alias_to_full[namemap[i,:别名]]["unit_gas"] = namemap[i,:度电耗气]
        end
        param = Dict()
        plant = Dict()
        ordered_unit = []
        for this_plant in gas_plan
            k = this_plant["plant"]
            plant[k] = Dict("total_gas"=>parse(Int,this_plant["gas"]),"units"=>[])
            if length(this_plant["plan"]) == 2 #无特殊要求，默认不过夜
                for n in 1:parse(Int,this_plant["plan"][1])
                    param["$k#$(n)机"] = Dict()
                    param["$k#$(n)机"]["Pmax"] = alias_to_full[k]["Pmax"]
                    param["$k#$(n)机"]["Pmin"] = 0.6*param["$k#$(n)机"]["Pmax"]
                    param["$k#$(n)机"]["单耗"] = alias_to_full[k]["unit_gas"]
                    param["$k#$(n)机"]["是否过夜"] = 0
                    push!(plant[k]["units"],"$k#$(n)机")
                    push!(ordered_unit,"$k#$(n)机")
                end
            else
                n = 1
                for mc in eachmatch(r"(?P<number>\d{1})[机组](?P<method>[\u4e00-\u9fa5]*)",this_plant["plan"])
                    for i in 1:parse(Int,mc["number"])
                        param["$k#$(n)机"] = Dict()
                        param["$k#$(n)机"]["Pmax"] = alias_to_full[k]["Pmax"]
                        param["$k#$(n)机"]["Pmin"] = 0.6*param["$k#$(n)机"]["Pmax"]
                        param["$k#$(n)机"]["单耗"] = alias_to_full[k]["unit_gas"]
                        if mc["method"]=="过夜" || mc["method"]=="连续运行"
                            param["$k#$(n)机"]["是否过夜"] = 1
                        else
                            param["$k#$(n)机"]["是否过夜"] = 0
                        end
                        push!(plant[k]["units"],"$k#$(n)机")
                        push!(ordered_unit,"$k#$(n)机")
                        n += 1
                    end
                end
            end    
        end
        nuclear,wind,solar,ext,load = next_day_schedule[!,:核电],next_day_schedule[!,:风电],next_day_schedule[!,:光伏],next_day_schedule[!,:受电],next_day_schedule[!,:统调负荷]
        gas_power = 0.95*next_day_schedule[!,:燃气]
        for list in (nuclear,wind,solar,ext,load,gas_power)
            tmp = [x for x in list]
            for i in 1:length(list)
                list[i] = tmp[(i+4*OFFSET-2)%96+1]
            end
        end
        COAL_PMAX,COAL_PMIN = config["统调燃煤机组最大发电能力"],config["统调燃煤机组最小发电能力"]
        RESERVE = config["应留备用"]
        WATER = config["水电发电能力"]
        m = Model(CPLEX.Optimizer)
        ### 默认计划分析....
        println("正在进行默认计划分析....")
        default_reserve = (COAL_PMAX + WATER) .+  nuclear + wind + solar + ext + gas_power - load
        if minimum(default_reserve) <= 2090
            println("应留备用为$RESERVE MW")
            println("默认计划存在备用缺口$(-minimum(default_reserve) + RESERVE)MW")
        else
            println("默认计划不存在备用缺口!备用最低值为$(minimum(default_reserve))MW")
        end 
        if minimum(default_reserve) < 0
            println("默认计划存在供电缺口$(-minimum(default_reserve))MW")
        end
        println("正在进行备用优化空间分析...") 
        ### 
        T = 96
        @variable(m,st[x in keys(param),t in 1:T],binary=true)
        @variable(m,up[x in keys(param),t in 1:T],binary=true)
        @variable(m,down[x in keys(param),t in 1:T],binary=true)
        @variable(m,hot[x in keys(param),t in 1:T],binary=true)
        @variable(m,cold[x in keys(param),t in 1:T],binary=true)
        # @variable(m,UT[x in keys(param),t in 0:T])
        # @variable(m,DT[x in keys(param),t in 0:T])
        @variable(m,Pg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,PMg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,dPg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,loadcut[t in 1:T],lower_bound=0)#正备用缺口
        @variable(m,windcut[t in 1:T],lower_bound=0)#负备用缺口
        @variable(m,rho)
        @variable(m,reserve[t in 1:T])
        @variable(m,neg_reserve[t in 1:T])
        @variable(m,dg[p in keys(plant)],lower_bound=0)
        # @variable(m,coal[t in 1:T],lower_bound=COAL_PMIN,upper_bound=COAL_PMAX)
        # @variable(m,water[t in 1:T],lower_bound=0,upper_bound=WATER)
        # @variable(m,fuel_cost[t in 1:T],lower_bound=0)
        for t in 1:T
            ### 机组约束
            for x in keys(param)
                @constraint(m,Pg[x,t] <= (1-cold[x,t])*param[x]["Pmax"])# 技术出力上限
                @constraint(m,Pg[x,t] >= hot[x,t]*param[x]["Pmin"])# 技术出力下限
                # @constraint(m,PMg[x,t] == (1 - hot[x,t]) * Pg[x,t] + hot[x,t]*param[x]["Pmax"])
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
                if t >= 2 #爬坡率约束
                    @constraint(m,Pg[x,t] - Pg[x,t-1] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                    @constraint(m,Pg[x,t-1] - Pg[x,t] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                else
                    if param[x]["是否过夜"] == 0
                        @constraint(m,Pg[x,t] <= hot[x,t]*param[x]["Pmax"]/3 + (1-hot[x,t])*param[x]["Pmin"]/4)
                    end
                end
                if t >= 2 #爬坡成本约束
                    @constraint(m,dPg[x,t] >= Pg[x,t] - Pg[x,t-1])
                    @constraint(m,dPg[x,t] >= Pg[x,t-1] - Pg[x,t])
                end
                if t >= 2
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - st[x,t-1])
                else
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - param[x]["是否过夜"])
                end
                @constraint(m,up[x,t] + down[x,t] <= 1)
                @constraint(m,up[x,t] <= 1 - sum(down[x,t-s] for s in 0:min(t-1,16)))#停机后间隔四个小时才能开机
                @constraint(m,down[x,t] <= 1 - sum(up[x,t-s] for s in 0:min(t-1,32)))#开机后间隔八个小时才能停机
                @constraint(m, hot[x,t] == st[x,t] - sum(up[x,t-s] for s in 0:min(t-1,3)))#启动好了
                @constraint(m, cold[x,t] == 1 - st[x,t] - sum(down[x,t-s] for s in 0:min(t-1,3)))#停好了
            end
            ### 机组约束
            # ## 煤机总体约束
            # for k in 0:20
            #     @constraint(m,fuel_cost[t] >= 0.25 * 100* (1 + k/20) * (coal[t] - (k/20 * COAL_PMAX + (20-k)/20 * COAL_PMIN)))
            # end
            ## 电网约束
            # @constraint(m,coal[t] + water[t] + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            #  sum(Pg[x,t] for x in keys(param)) == load[t] - loadcut[t])
            @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(PMg[x,t] for x in keys(param)) >= load[t] + RESERVE - loadcut[t]) #正备用约束
            @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(PMg[x,t] for x in keys(param)) == load[t] + RESERVE + reserve[t]) #正备用约束
            @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(Pg[x,t] for x in keys(param)) <= load[t] + windcut[t]) #负备用约束
            @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(Pg[x,t] for x in keys(param)) == load[t] - neg_reserve[t]) #负备用约束
            @constraint(m,rho <= reserve[t])
        end
        for x in keys(param)
            if param[x]["是否过夜"] == 1
                @constraint(m,sum(up[x,t] for t in 1:T) == 0)
                @constraint(m,sum(down[x,t] for t in 1:T) == 0)
            else
                @constraint(m,sum(up[x,t] for t in 1:T) == 1)
                @constraint(m,sum(down[x,t] for t in 1:T) == 1)
                for t in T-20:T
                    @constraint(m,st[x,t] == 0)
                end
            end
        end
        for p in keys(plant) #总气量约束
            @constraint(m,0.1*sum(Pg[x,t]*0.25*param[x]["单耗"] for x in plant[p]["units"] for t in 1:T) + dg[p] == plant[p]["total_gas"])
        end
        @objective(m,Min,10000*0.25*sum(loadcut) + sum(windcut) - rho + 1000*sum(dg))
        # @objective(m,Min,sum(((1-cold[x,t])*param[x]["Pmax"] - Pg[x,t]) for x in keys(param) for t in 1:T))
        set_time_limit_sec(m, config["maxTimeformaxReserve"])
        if config["logmaxReserve"]
            optimize!(m)
        else
            @suppress_out optimize!(m)
        end
        println("备用最低值可优化至$(round(value(rho)) + RESERVE)MW\n")
        df = DataFrame()
        df[!,"时刻"] = [next_day + Hour(OFFSET) + Minute((t-1)*15) for t in 1:T]
        for x in ordered_unit
            if param[x]["是否过夜"] == 1
                df[!,x] = [1 for t in 1:T]
            else
                df[!,x] = [missing for t in 1:T]
            end
        end
        XLSX.writetable("assets/必开燃机列表.xlsx",df;overwrite=true)
        if Sys.iswindows()
            run(`cmd /c .\\assets\\必开燃机列表.xlsx`)
        else
            run(`open ./assets/必开燃机列表.xlsx`)
        end
        return 0# if things finished successfully    
    catch err
        showerror(stdout, err, catch_backtrace())
        println("press Enter to quit...")
        s = readline()
        print(s)
        return 0
    end
  end
function UC()::Cint
    try
        config = JSON.parsefile("assets/config.json")
        OFFSET = 7 #从第二天的7点开始规划
        function myfun(r)
            v = r[:时间]
            return v == "高峰电量"
        end
        # filename = "assets/日前计划编制报表20220721.xlsx"
        println("请选择日前计划文件（仅支持xlsx格式）！")
        filename = pick_file("";filterlist="xlsx")
        next_day_schedule = DataFrame(XLSX.readtable(filename,1,"A:EG";
        first_row=3,stop_in_row_function=myfun)...)
        required_uc = DataFrame(XLSX.readtable("assets/必开燃机列表.xlsx",1)...)
        next_day_string = replace(split(XLSX.readdata(filename,1,"AJ2"))[1],r"[年月日]"=>"")
        next_day = DateTime(next_day_string,DateFormat("yyyymmdd"))
        row_string = XLSX.readtable(filename, 1,"A:EG";
        first_row=3)[1][2][end]
        re = r"燃机开机容量.*万方"
        gas_plan_string = collect(eachmatch(re,row_string))[1].match
        total_capacity = match(r"燃机开机容量(?P<total>\d*)万千瓦",gas_plan_string)["total"]
        # gas_plan = collect(eachmatch(r"[\u4e00-\u9fa5]*[1-9]{1}[机组]",gas_plan_string))
        gas_plan = collect(eachmatch(r"(?P<plant>[\u4e00-\u9fa5]*)(?P<plan>[1-9]{1}[机组][^，]*)（(?P<gas>\d*)万方",gas_plan_string))
        # gas_plan_tmp = [[c for c in x.match] for x in gas_plan]
        # gas_plan_unpack = Array([(join(x[1:end-2],""),parse(Int,x[end-1])) for x in gas_plan_tmp])
        namemap = DataFrame(XLSX.readtable("assets/燃机名称表.xlsx","Sheet1","A:F")...)
        # gas_total = []
        # for s in eachmatch(r"[\(（][1-9]\d*万方",gas_plan_string)
        #     push!(gas_total,parse(Int,match(r"[1-9]\d*",s.match).match))
        # end
        alias_to_full = Dict()
        for i in range(1,stop=nrow(namemap))
            alias_to_full[namemap[i,:别名]] = Dict()
            alias_to_full[namemap[i,:别名]]["Pmax"] = namemap[i,:单机出力上限]
            alias_to_full[namemap[i,:别名]]["unit_gas"] = namemap[i,:度电耗气]
            alias_to_full[namemap[i,:别名]]["输出名称"] = namemap[i,:输出名称]
        end
        param = Dict()
        plant = Dict()
        ordered_unit = []
        for this_plant in gas_plan
            k = this_plant["plant"]
            plant[k] = Dict("total_gas"=>parse(Int,this_plant["gas"]),"units"=>[])
            if length(this_plant["plan"]) == 2 #无特殊要求，默认不过夜
                for n in 1:parse(Int,this_plant["plan"][1])
                    param["$k#$(n)机"] = Dict()
                    param["$k#$(n)机"]["Pmax"] = alias_to_full[k]["Pmax"]
                    param["$k#$(n)机"]["Pmin"] = 0.6*param["$k#$(n)机"]["Pmax"]
                    param["$k#$(n)机"]["单耗"] = alias_to_full[k]["unit_gas"]
                    param["$k#$(n)机"]["是否过夜"] = 0
                    param["$k#$(n)机"]["输出名称"] = alias_to_full[k]["输出名称"] * "#$(n)机"
                    push!(plant[k]["units"],"$k#$(n)机")
                    push!(ordered_unit,"$k#$(n)机")
                end
            else
                n = 1
                for mc in eachmatch(r"(?P<number>\d{1})[机组](?P<method>[\u4e00-\u9fa5]*)",this_plant["plan"])
                    for i in 1:parse(Int,mc["number"])
                        param["$k#$(n)机"] = Dict()
                        param["$k#$(n)机"]["Pmax"] = alias_to_full[k]["Pmax"]
                        param["$k#$(n)机"]["Pmin"] = 0.6*param["$k#$(n)机"]["Pmax"]
                        param["$k#$(n)机"]["单耗"] = alias_to_full[k]["unit_gas"]
                        param["$k#$(n)机"]["输出名称"] = alias_to_full[k]["输出名称"] * "#$(n)机"
                        if mc["method"]=="过夜" || mc["method"]=="连续运行"
                            param["$k#$(n)机"]["是否过夜"] = 1
                        else
                            param["$k#$(n)机"]["是否过夜"] = 0
                        end
                        push!(plant[k]["units"],"$k#$(n)机")
                        push!(ordered_unit,"$k#$(n)机")
                        n += 1
                    end
                end
            end    
        end
        ## 必开燃机中过夜
        for x in keys(param)
            if !ismissing(sum(required_uc[!,x]))
                if sum(required_uc[!,x]) == 96
                    param[x]["是否过夜"] = 1
                end
            end
        end
        nuclear,wind,solar,ext,load = next_day_schedule[!,:核电],next_day_schedule[!,:风电],next_day_schedule[!,:光伏],next_day_schedule[!,:受电],next_day_schedule[!,:统调负荷]
        gas_power = next_day_schedule[!,:燃气]
        for list in (nuclear,wind,solar,ext,load,gas_power)
            tmp = [x for x in list]
            for i in 1:length(list)
                list[i] = tmp[(i+4*OFFSET-2)%96+1]
            end
        end
        COAL_PMAX,COAL_PMIN = config["统调燃煤机组最大发电能力"],config["统调燃煤机组最小发电能力"]
        RESERVE = config["期望备用"]
        WATER = config["水电发电能力"]
        m = Model(CPLEX.Optimizer)
        println("正在进行备用优化...") 
        ### 
        T = 96
        @variable(m,st[x in keys(param),t in 1:T],binary=true)
        @variable(m,up[x in keys(param),t in 1:T],binary=true)
        @variable(m,down[x in keys(param),t in 1:T],binary=true)
        @variable(m,hot[x in keys(param),t in 1:T],binary=true)
        @variable(m,cold[x in keys(param),t in 1:T],binary=true)
        # @variable(m,UT[x in keys(param),t in 0:T])
        # @variable(m,DT[x in keys(param),t in 0:T])
        @variable(m,Pg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,dPg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,PMg[x in keys(param),t in 1:T],lower_bound=0)
        @variable(m,loadcut[t in 1:T],lower_bound=0)#正备用缺口
        @variable(m,windcut[t in 1:T],lower_bound=0)#负备用缺口
        @variable(m,rho)
        @variable(m,reserve[t in 1:T])
        @variable(m,neg_reserve[t in 1:T])
        @variable(m,dg[p in keys(plant)],lower_bound=0)
        # @variable(m,coal[t in 1:T],lower_bound=COAL_PMIN,upper_bound=COAL_PMAX)
        # @variable(m,water[t in 1:T],lower_bound=0,upper_bound=WATER)
        # @variable(m,fuel_cost[t in 1:T],lower_bound=0)
        gas_cons_ref = []
        for t in 1:T
            ### 机组约束
            for x in keys(param)
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
                if t >= 2 #爬坡率约束
                    @constraint(m,Pg[x,t] - Pg[x,t-1] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                    @constraint(m,Pg[x,t-1] - Pg[x,t] <= hot[x,t-1]*param[x]["Pmax"]/3 + (1-hot[x,t-1])*param[x]["Pmin"]/4)
                else
                    if param[x]["是否过夜"] == 0
                        @constraint(m,Pg[x,t] <= hot[x,t]*param[x]["Pmax"]/3 + (1-hot[x,t])*param[x]["Pmin"]/4)
                    end
                end
                if t >= 2 #爬坡成本约束
                    @constraint(m,dPg[x,t] >= Pg[x,t] - Pg[x,t-1])
                    @constraint(m,dPg[x,t] >= Pg[x,t-1] - Pg[x,t])
                end
                if t >= 2
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - st[x,t-1])
                else
                    @constraint(m,up[x,t] - down[x,t] == st[x,t] - param[x]["是否过夜"])
                end
                @constraint(m,up[x,t] + down[x,t] <= 1)
                @constraint(m,up[x,t] <= 1 - sum(down[x,t-s] for s in 0:min(t-1,16)))#停机后间隔四个小时才能开机
                @constraint(m,down[x,t] <= 1 - sum(up[x,t-s] for s in 0:min(t-1,32)))#开机后间隔四个小时才能停机
                @constraint(m, hot[x,t] == st[x,t] - sum(up[x,t-s] for s in 0:min(t-1,3)))#启动好了
                @constraint(m, cold[x,t] == 1 - st[x,t] - sum(down[x,t-s] for s in 0:min(t-1,3)))#停好了
            end
            ### 机组约束
            # ## 煤机总体约束
            # for k in 0:20
            #     @constraint(m,fuel_cost[t] >= 0.25 * 100* (1 + k/20) * (coal[t] - (k/20 * COAL_PMAX + (20-k)/20 * COAL_PMIN)))
            # end
            ## 电网约束
            # @constraint(m,coal[t] + water[t] + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            #  sum(Pg[x,t] for x in keys(param)) == load[t] - loadcut[t])
            @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(PMg[x,t] for x in keys(param)) >= load[t] + RESERVE - loadcut[t]) #正备用约束
            @constraint(m,COAL_PMAX + WATER + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(PMg[x,t] for x in keys(param)) == load[t] + RESERVE + reserve[t]) #正备用约束
            @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(Pg[x,t] for x in keys(param)) <= load[t] + windcut[t]) #负备用约束
            @constraint(m,COAL_PMIN + sum(l[t] for l in [nuclear,wind,solar,ext]) +
            sum(Pg[x,t] for x in keys(param)) == load[t] - neg_reserve[t]) #负备用约束
            @constraint(m,rho <= reserve[t])
        end
        for x in keys(param)
            if param[x]["是否过夜"] == 1
                @constraint(m,sum(up[x,t] for t in 1:T) == 0)
                @constraint(m,sum(down[x,t] for t in 1:T) == 0)
            else
                @constraint(m,sum(up[x,t] for t in 1:T) <= config["最多启停次数"])
                @constraint(m,sum(down[x,t] for t in 1:T) <= config["最多启停次数"])
                for t in 1:T-19
                    if !ismissing(required_uc[!,x][t])
                        @constraint(m,st[x,t] == 1)
                    end
                end
                for t in T-20:T
                    @constraint(m,st[x,t] == 0)
                end
            end
        end
        for p in keys(plant) #总气量约束
            con = @constraint(m,0.1*sum(Pg[x,t]*0.25*param[x]["单耗"] for x in plant[p]["units"] for t in 1:T) + dg[p] == plant[p]["total_gas"])
            push!(gas_cons_ref,con)
        end
        @objective(m,Min,10000*0.25*sum(loadcut) + sum(windcut) + 0.1*sum(dPg) + 100*sum(st) + 10000*sum(dg))
        set_time_limit_sec(m, config["maxTimeforUC"])
        if config["logUC"]
            optimize!(m)
        else
            @suppress_out optimize!(m)
        end
        TT = [(t+4*(24-OFFSET)-2)%96+1 for t in 1:T]
        df = DataFrame()
        df[!,"时刻"] = [next_day + Minute(t*15) for t in 1:T]
        df[!,"煤机最大发电能力"] = [COAL_PMAX for t in TT]
        df[!,"煤机最小发电能力"] = [COAL_PMIN for t in TT]
        df[!,"燃机"] = Int.(round.([sum(value(Pg[x,t]) for x in keys(param)) for t in TT]))
        df[!,"水电"] = [WATER for t in TT]
        df[!,"核电"] = [nuclear[t] for t in TT]
        df[!,"风电"] = [wind[t] for t in TT]
        df[!,"光伏"] = [solar[t] for t in TT]
        df[!,"受电"] = [ext[t] for t in TT]
        df[!,"统调负荷"] = [load[t] for t in TT]
        df[!,"正备用"] = Int.(round.([value(reserve[t]) + RESERVE for t in TT]))
        df[!,"负备用"] = Int.(round.([value(neg_reserve[t]) for t in TT]))
        for x in ordered_unit
            df[!,x] = Int.(round.([value(Pg[x,t]) for t in TT]))
        end
        df2 = DataFrame()
        df2[!,"燃气电厂"] = [alias_to_full[p]["输出名称"] for p in keys(plant)]
        df2[!,"计划气量"] = [plant[p]["total_gas"] for p in keys(plant)]
        df3 = DataFrame()
        df3[!,"时刻"] = [Dates.format(next_day + Minute(t*15),dateformat"yyyy-mm-dd HH:MM:SS") for t in 1:T]
        for x in ordered_unit
            df3[!,param[x]["输出名称"]] = Int.(round.([value(Pg[x,t]) for t in TT]))
        end
        println("优化成功！")
        println("备用最低值为$(minimum(value.(reserve)) + RESERVE)")
        # plants_keys = collect(keys(plant))
        # for i in 1:length(plants_keys)
        #     println("$(plants_keys[i]) 的气量影子价格为:$(dual(gas_cons_ref[i]))")
        # end
        save_file_name = save_file("";filterlist="xlsx")
        # save_file_name = save_dialog("结果保存至...", GtkNullContainer(), (GtkFileFilter("*.xlsx", name="All supported formats"), "*.xlsx"))
        # XLSX.writetable(save_file_name,df;overwrite=true)
        XLSX.writetable(save_file_name,"REPORT_A"=>df,"REPORT_B"=>df2,"REPORT_C"=>permutedims(df3,"时刻");overwrite=true)
        lowest,p = findmin(df[!,"正备用"][1:68])
        tl,tc,tr,te,tso,twi = round(df[!,"统调负荷"][p]/10),round(df[!,"煤机最大发电能力"][p]/10),round(df[!,"燃机"][p]/10),round(df[!,"受电"][p]/10),round(df[!,"光伏"][p]/10),round(df[!,"风电"][p]/10)
        lowest_time = Dates.format(df[!,"时刻"][p],"HH:MM")
        println("明日备用情况：按计划曲线考虑，$(lowest_time)备用最紧，为$(Int(round(lowest/10)))万千瓦；该时刻，统调负荷$(Int(tl))万千瓦，煤机$(Int(tc))万千瓦，燃机$(Int(tr))万千瓦，受电$(Int(te))万千瓦，统调光伏$(Int(tso))万千瓦，统调风电$(Int(twi))万千瓦。")
        println("明日燃机开机容量$(total_capacity)万千瓦，总气量$(Int(sum(df2[!,"计划气量"])))万方。")
        for this_plant in gas_plan
            k = this_plant["plant"]
            tmpstr = ""
            gy,ngy = sum(param[x]["是否过夜"] == 1 for x in plant[k]["units"]),sum(param[x]["是否过夜"] == 0 for x in plant[k]["units"])
            if ngy > 0 
                tmpstr *= "$(ngy)机日开夜停"
            end
            if gy > 0
                tmpstr *= "$(gy)机过夜"
            end
            print("$k $(tmpstr)（$(plant[k]["total_gas"])万方）。")
            if length(plant[k]["units"]) == 1
                x = plant[k]["units"][1]
                if param[x]["是否过夜"] == 1
                    load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
                    if load_rate <= 90
                        print("（需压出力运行）\n")
                    else
                        print("\n")
                    end
                else
                    start,stop = findfirst([value(hot[x,t])>0 for t in 1:T]),findlast(([value(hot[x,t])>0 for t in 1:T]))
                    start,stop = df[!,"时刻"][start],df[!,"时刻"][stop]
                    load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
                    if hour(stop) > 8
                        print("$(Dates.format(start,"HH:MM"))带足--$(Dates.format(stop,"HH:MM"))")
                    else
                        print("$(Dates.format(start,"HH:MM"))带足--次日$(Dates.format(stop,"HH:MM"))")
                    end
                    if load_rate <= 95
                        print("（需压出力运行）\n")
                    else
                        print("\n")
                    end
                end
            else
                print("\n")
                for i in 1:length(plant[k]["units"])
                    x = plant[k]["units"][i]
                    if param[x]["是否过夜"] == 1
                        load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
                        print("     #$(i)机过夜")
                        if load_rate <= 90
                            print("（需压出力运行）\n")
                        else
                            print("\n")
                        end
                    else
                        start,stop = findfirst([value(hot[x,t])>0 for t in 1:T]),findlast(([value(hot[x,t])>0 for t in 1:T]))
                        start,stop = df[!,"时刻"][start],df[!,"时刻"][stop]
                        load_rate = round(100*sum(value(hot[x,t])*value(Pg[x,t]) for t in 1:T)/(sum(value.(hot[x,:]))*param[x]["Pmax"]))
                        if hour(stop) > 8
                            print("     #$(i)机: $(Dates.format(start,"HH:MM"))带足--$(Dates.format(stop,"HH:MM"))")
                        else
                            print("     #$(i)机: $(Dates.format(start,"HH:MM"))带足--次日$(Dates.format(stop,"HH:MM"))")
                        end
                        if load_rate <= 90
                            print("（需压出力运行）\n")
                        else
                            print("\n")
                        end
                    end
                end
            end
        end
        println("press Enter to quit...")
        s = readline()
        print(s)
        return 0# if things finished successfully
    catch err
        showerror(stdout, err, catch_backtrace())
        println("press Enter to quit...")
        s = readline()
        print(s)
        return 0
    end
end
end # module