#' extract_info_from_xlsx
#'
#' @description This function opens a Excel file and outputs the GEM Type
#' and File Locations from the "Info" Sheet in a list.
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
#'
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#'
#' @return Returns a list with THGEM type (list$gem_type) and the location of measurement files
#'  (list$file_locations)
#' @export
#'
#' @importFrom openxlsx loadWorkbook
#' @importFrom openxlsx read.xlsx
#'
#' @examples \dontrun{
#'     data_info <- extract_info_from_xlsx(excel_file = "Data_Analysis.xlsx")
#'
#'     cat("THGEM Type: ",data_info$gem_type)
#'     cat("File Locations: ",data_info$file_locations)
#' }

extract_info_from_xlsx <- function(excel_file){

  File <- excel_file
  sheet <- "Info"
  wb <- openxlsx::loadWorkbook(file = File)

  #Read a section of a sheet from the Excel File for general Information
  Data_Info <- openxlsx::read.xlsx(wb, sheet = "Info",
                                   colNames = FALSE, #Column Names are there
                                   rows = seq(1:4), #Copy out these rows
                                   cols = seq(1:5)) #Copy out these columns
  gem_type = Data_Info[1,2]
  #Extract Measurement Folder Information and define filepath to said folder

  #Define Location of measurement data files
  file_locations <- gsub("\\\\","/",Data_Info[2,2]) #gsub needs \\ for \.
  file_locations <- paste(file_locations, gem_type, sep="")


  results <- list("gem_type" = gem_type, "file_locations" = file_locations)
  cat("\n THGEM Type: ", gem_type, "\n File Locations: ", file_locations)
  return(results)

}
