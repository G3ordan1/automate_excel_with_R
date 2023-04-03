library(openxlsx)
library(tidyverse)
library(tidyquant)
library(timetk)


# Get stock data of apple, google, netflix, nvidia and microsoft for the period of 5000 days.
stk_data_table <- c("AAPL", "GOOG", "NFLX", "NVDA", "MSFT") %>%
  tq_get(from = Sys.Date() - 5000, to = Sys.Date())

view(stk_data_table)

# Pivot table of stock data with year as rows and symbol as columns.
stk_pivot_table <- stk_data_table %>%
  pivot_table(
    .rows = ~ YEAR(date),
    .column = ~symbol,
    .values = ~ PCT_CHANGE_FIRSTLAST(adjusted)
  ) %>%
  rename(year = 1) #rename the first column to year

# Pivot table of stock data with date as rows and symbol as columns.
stk_pivot_table_adjusted <- stk_data_table %>%
  pivot_table(
    .rows = ~date,
    .column = ~symbol,
    .values = ~adjusted
  )

view(stk_pivot_table_adjusted)

# Plot of stock data grouped by symbol and with date as x axis and adjusted as y axis.
stock_plot <- stk_data_table %>%
  group_by(symbol) %>%
  plot_time_series(
    date,
    adjusted,
    .color_var = symbol,
    .facet_ncol = 2,
    .interactive = FALSE
  )


wb <- createWorkbook()
addWorksheet(wb, sheetName = "stock_data")
addWorksheet(wb, sheetName = "stock_plot")
print(stock_plot)
insertPlot(wb, sheet = "stock_plot", startRow = 2, startCol = 2)
writeDataTable(wb, sheet = "stock_data", x = stk_pivot_table)
saveWorkbook(wb, "stocks.xlsx", overwrite = TRUE)
