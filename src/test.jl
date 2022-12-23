using Revise
using UnicodePlots
includet("myUCApp.jl")
using .myUCApp
# maxReserve()
UC()
# ## 会编译报错的package: Plots
# using XLSX
# XLSX.openxlsx("assets/次日平衡模板.xlsx",mode="rw") do xf
#     sheet = xf[1]
#     sheet[1:5,20] = 1111
# end