#' save_dataFrame_back_to_xlsx
#'
#' @description Saves dataframe back into Excel Sheet.
#'
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#' @param excel_sheet_name Name of Excel Sheet
#' @param excel_data_frame Dataframe, which should be saved back to Excel file. Obtained using
#' \code{\link{extract_dataFrame_from_xlsx}}.
#'
#' @return Executed for its side effects.
#'
#' @export
#'
#' @importFrom openxlsx writeData
#' @importFrom openxlsx loadWorkbook
#' #' @importFrom openxlsx saveWorkbook
#' @examples
save_dataFrame_back_to_xlsx <- function(excel_file,
                                        excel_sheet_name,
                                        excel_data_frame){

  File <- excel_file
  sheet <- excel_sheet_name
  wb2 <- openxlsx::loadWorkbook(file = File)
  DataFromSource <- excel_data_frame

  openxlsx::writeData(wb2, sheet = sheet, x = DataFromSource,
            startCol = "A",
            startRow = 2,
            colNames = FALSE) # Do not print Header Line with Column Names
  openxlsx::saveWorkbook(wb2, File, overwrite = TRUE)

  cat("\nData has been saved back to Excel file.")


 return()
}
