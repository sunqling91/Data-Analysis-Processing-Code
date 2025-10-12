# 安装并加载必要的包
library(metafor)
library(readxl)

# 读取Excel文件中的数据
file_path <- "C:/Users/lenovo/Desktop/计算贡献25_1/偏相关结果_2.5.xlsx"
data <- read_excel(file_path, sheet = "Sheet1")

# 进行元分析
res <- rma(yi = data$Slope, sei = data$Std, data = data,method = "DL")  # 或 method = "DL"


# 输出结果s
print(res)

