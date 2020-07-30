#' @title extract_dataFrame_from_xlsx
#'
#' @description This function opens a certain sheet of an Excel file (.xlsx)
#'   and copies out a select number of rows and columns into a R dataframe.
#'
#' @details Excel sheet must contain a sheet named "Info", where in cell B1 the
#'   THGEM type of the measurements is given and in cell B2 the path to the
#'   folder with measurement data is given. In the general measurement data folder,
#'   subfolders for each THGEM should be present with the THGEM name as folder name.
#'
#'   Folder structure:
#'   \itemize{
#'   \item{.../Measurement_Data}
#'     \itemize{
#'     \item{em1mm}
#'     \item{em1p5mm}
#'     \item{em2mm}
#'     \item{keramik1p5mm}
#'   }}
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#' @param excel_sheet_name Name of Excel Sheet
#' @param excel_nrows Number of rows to be copied out, after header (1:(excel_nrows+1)). Default: 100.
#' @param excel_ncols Number of columns to be copied out (1:excel_ncols). Default: 180.
#' @return data.frame
#' @export
#' @importFrom openxlsx loadWorkbook
#' @importFrom openxlsx read.xlsx
#' @examples \dontrun{
#' DataFromExcel <- extract_dataFrame_from_xlsx(excel_file = "Data_Analysis_Mobility.xlsx",
#'                    excel_sheet_name = "M1",
#'                    excel_nrows = 100,
#'                    excel_ncols = 180)
#' }
extract_dataFrame_from_xlsx <- function(excel_file,
                               excel_sheet_name,
                               excel_nrows = 100,
                               excel_ncols = 180){
  File <- excel_file
  sheet <- excel_sheet_name
  wb <- openxlsx::loadWorkbook(file = File)


  #Define number of rows from Excel to be imported
  ExcelRows = c(1,2:excel_nrows) #c(1,2:X) -> Header = 1, rest start at row 2, end at Meas X-1

  #Import Section from Excel File into data frame
  DataFromSource <- openxlsx::read.xlsx(wb, sheet = sheet,
                                        colNames = TRUE, #Column Names are there
                                        rows = ExcelRows, #Copy out these rows
                                        cols = seq(1:excel_ncols)) #Copy out these columns

  cat("\n--- Data extraction from Excel Sheet was succesful-----------------------------------------")
  return(DataFromSource)
}
