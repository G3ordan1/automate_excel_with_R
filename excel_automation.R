library(openxlsx)
library(tidyverse)
library(tidyquant)
library(timetk)

stk_data_table <- c("AAPL", "GOOG", "NFLX", "NVDA", "MSFT") %>%
  tq_get(from = Sys.Date() - 5000, to = Sys.Date())

view(stk_data_table)

stk_pivot_table <- stk_data_table %>%
  pivot_table(.rows = ~ YEAR(date), .column = ~symbol, .values = ~ PCT_CHANGE_FIRSTLAST(adjusted)) %>%
  rename(year = 1)

stk_pivot_table_adjusted  <- stk_data_table %>%
  pivot_table(
  .rows = ~ date,
  .column = ~ symbol,
  .values = ~ adjusted
)

view(stk_pivot_table_adjusted)

stock_plot <- stk_data_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, .color_var = symbol, .facet_ncol = 2, .interactive = FALSE)


wb  <- createWorkbook()
addWorksheet(wb, sheetName = "Stock_data")
addWorksheet(wb, sheetName = "stock_plot")
print(stock_plot)
insertPlot(wb, sheet = "stock_plot", startRow = 1, startCol = 1)
writeDataTable(wb, sheet = "Stock_data", x = stk_pivot_table)
saveWorkbook(wb, "stocks.xlsx", overwrite = TRUE)
