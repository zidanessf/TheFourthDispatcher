using Revise
# using UnicodePlots
includet("myUCApp.jl")
using .myUCApp
# maxReserve()
UC()
# ## 会编译报错的package: Plots
# using XLSX
# XLSX.openxlsx("assets/次日平衡模板备份.xlsx",mode="rw") do xf
#     sheet = xf[1]
#     for i in 2:18
#         sheet[i,2] = 2222
#     end
# end