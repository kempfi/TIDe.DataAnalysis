#' data_analysis_filename
#' @description Extracts the filename from a dataframe and then splits it at "_" and pastes the different
#' parameters into the correct excel columns. Only executed if Data$Filename.Read.in[i]==1.
#' @param excel_data_frame Dataframe of excel sheet to be analysed, obtained with \code{\link{extract_dataFrame_from_xlsx}}.
#'
#' @return Dataframe with file name extraction completed
#' @export
#'
#' @examples \dontrun{
#' DataFromExcel <- extract_dataFrame_from_xlsx(excel_file = "Data_Analysis_Mobility.xlsx",
#'                    excel_sheet_name = "M1",
#'                    excel_nrows = 100,
#'                    excel_ncols = 180)
#' DataFromExcel_with_filenames <- data_analysis_filename(data = DataFromExcel)
#' }
data_analysis_filename <- function(excel_data_frame)
  {

  DataFromSource <- excel_data_frame
  excel_nrows <- dim(excel_data_frame)[1]

  #Split Filename
  filename_split = strsplit(DataFromSource$Filename..txt, "_")


  #Enter Data from splitted Filename. First check Column "Read in Filename".
  #If 0 do not overwrite. If 1 do overwrite with file name info
  for( i in 1:(excel_nrows)){       # -1 because header is not included for filename writing
    cat("\nMeasurement: ", i)
    if (is.na(DataFromSource$Filename.Read.in[i])){
      DataFromSource$Filename.Read.in[i] <- 0
    }

    if (DataFromSource$Filename.Read.in[i]==1){
      #write in GEM Type
      DataFromSource[i,4] <- filename_split[[i]][1] #GEM Type

      #Convert pressure into number (p ->.) and save
      DataFromSource[i,5] <- as.numeric(gsub("p",".",filename_split[[i]][2])) #Druck

      #Write in HV
      DataFromSource[i,6] <- as.numeric(gsub("V","",filename_split[[i]][4])) #HV

      #Write in Drift after removing V
      DataFromSource[i,7] <- as.numeric(gsub("V","",filename_split[[i]][5])) #Drift

      #Write time in ms for window
      DataFromSource[i,8] <- as.numeric(gsub("ms","",filename_split[[i]][6])) #t

      #Write in Amplification stats
      DataFromSource[i,9] <- filename_split[[i]][7] #Amplification

      #Write in Threshold input
      DataFromSource[i,11] <- filename_split[[i]][8] #Threshold input

      #Write in Threshold
      DataFromSource[i,10] <- as.numeric(gsub("p",".",gsub("mVthreshold","",filename_split[[i]][9]))) #Threshold in mV

      #Write in Additional Information
      DataFromSource[i,13] <- filename_split[[i]][10] #Additional Information and comments

      #Write in Date
      DataFromSource[i,12] <- as.numeric(filename_split[[i]][11]) #Date

      #Write in Meas No
      DataFromSource[i,14] <- as.numeric(gsub(".txt","",filename_split[[i]][12])) #Measurement number

      cat(" Filename Info written")
    }
    else if (DataFromSource$Filename.Read.in[i]==0){
      cat(" Filename Info skipped")
    }
  }

  return(DataFromSource)
}
