# 安装并加载必要的包
library(metafor)
library(readxl)

# 读取Excel文件中的数据
file_path <- "C:/Users/lenovo/Desktop/土壤站点数据/11-14/环境变量线性回归趋势.xlsx"
data <- read_excel(file_path, sheet = "meta")

# 进行元分析
res <- rma(yi = data$Slope, sei = data$Std, data = data,method = "DL")

# 输出结果
print(res)
