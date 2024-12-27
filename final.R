# 安裝必要套件 
install.packages("tidyverse")
install.packages("openxlsx")
# 載入必要的套件
library(tidyverse)
library(openxlsx)

# 讀取資料，檔案名稱為 input.csv
input_data <- read.csv("input.csv", fileEncoding = "UTF-8")

# 確認列名
print(colnames(input_data))

# 按年度加總進口總值和出口總值，並換算為萬元
data_summarized <- input_data %>%
  mutate(年度 = as.integer(年度)) %>%  # 將年度轉為整數
  group_by(年度) %>%
  summarize(
    年進口總值 = sum(`進口總值.新臺幣千元.`, na.rm = TRUE) / 10,  # 換算為萬元
    年出口總值 = sum(`出口總值.新臺幣千元.`, na.rm = TRUE) / 10   # 換算為萬元
  )

# 確認加總後的結果
print(data_summarized)

# 視覺化：進出口總值的年度趨勢
plot <- ggplot(data_summarized, aes(x = 年度)) +
  geom_line(aes(y = 年進口總值, color = "進口總值"), linewidth = 1) +
  geom_line(aes(y = 年出口總值, color = "出口總值"), linewidth = 1) +
  scale_x_continuous(breaks = seq(min(data_summarized$年度), max(data_summarized$年度), by = 1)) +  # 升降單位為 1
  labs(
    title = "進出口總值年度趨勢",
    x = "年度",
    y = "總值 (單位: 萬元)",
    color = "類別"
  ) +
  theme_classic()

# 創建 Excel 檔案並寫入數據和圖表
wb <- createWorkbook()

# 添加表格工作表
addWorksheet(wb, "年度匯總表")
writeData(wb, "年度匯總表", data_summarized)

# 添加圖表工作表
addWorksheet(wb, "進出口圖表")

# 將 ggplot 圖表保存為臨時文件
temp_file <- tempfile(fileext = ".png")
ggsave(temp_file, plot = plot, width = 8, height = 6)

# 將圖表插入到 Excel 中
insertImage(wb, "進出口圖表", temp_file, startRow = 1, startCol = 1)

# 保存 Excel 文件
saveWorkbook(wb, "進出口總值分析.xlsx", overwrite = TRUE)

cat("資料分析完成，結果已保存為 '進出口總值分析.xlsx'")
