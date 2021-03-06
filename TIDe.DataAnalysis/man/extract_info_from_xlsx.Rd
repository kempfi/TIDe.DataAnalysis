% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/extract_info_from_xlsx.R
\name{extract_info_from_xlsx}
\alias{extract_info_from_xlsx}
\title{extract_info_from_xlsx}
\usage{
extract_info_from_xlsx(excel_file)
}
\arguments{
\item{excel_file}{Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"}
}
\value{
Returns a list with THGEM type (list$gem_type) and the location of measurement files
 (list$file_locations)
}
\description{
This function opens a Excel file and outputs the GEM Type
and File Locations from the "Info" Sheet in a list.
}
\details{
Excel sheet must contain a sheet named "Info", where in cell B1 the
  THGEM type of the measurements is given and in cell B2 the path to the
  folder with measurement data is given. In the general measurement data folder,
  subfolders for each THGEM should be present with the THGEM name as folder name.

  Folder structure:
  \itemize{
  \item{.../Measurement_Data}
    \itemize{
    \item{em1mm}
    \item{em1p5mm}
    \item{em2mm}
    \item{keramik1p5mm}
  }}
}
\examples{
\dontrun{
    data_info <- extract_info_from_xlsx(excel_file = "Data_Analysis.xlsx")

    cat("THGEM Type: ",data_info$gem_type)
    cat("File Locations: ",data_info$file_locations)
}
}
