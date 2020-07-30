#' data_analysis_darkRate
#' @description Extracts the dark rate measurement data from a dataframe and then analyses it and saves the results
#' into the correct excel columns. Only executed if Data$Data.analysis[i]==1.
#'
#' @param excel_data_frame Dataframe of excel sheet to be analysed, obtained with \code{\link{extract_dataFrame_from_xlsx}}.
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#' @param alpha_packet_size Number of alpha particles to be grouped together into one data point,
#' default 100.
#' @importFrom("graphics", "hist", "plot")
#' @importFrom("dplyr", "filter")
#' @importFrom("stats", "sd")
#' @return Dataframe with dark rate analysis completed.
#' @export
#'
#' @examples \dontrun{
#' DataFromExcel <- extract_dataFrame_from_xlsx(excel_file = "Data_Analysis_Mobility.xlsx",
#'                    excel_sheet_name = "M1",
#'                    excel_nrows = 100,
#'                    excel_ncols = 180)
#' DataFromExcel <- data_analysis_filename(data = DataFromExcel)
#' DataFromExcel <- data_analysis_main(data = DataFromExcel,
#'                                     excel_file = "Data_Analysis_Mobility.xlsx")
#' }
data_analysis_darkRate <- function(excel_file,
                               excel_data_frame,
                               alpha_packet_size=100


){

  DataFromSource <- excel_data_frame
  excel_nrows <- dim(excel_data_frame)[1]

  for( i in 1:(excel_nrows)){       # -1 because header is not included for filename writing
    cat("\n--- Measurement ", i, "-------------------------------------------------------------------")

     if(is.na(DataFromSource$DR.Data.analysis[i])){
        DataFromSource$DR.Data.analysis[i] <- 0}

    if (DataFromSource$DR.Data.analysis[i]==1){ #Check if overwrite is activated
      cat("\nDR Data analysis of i=", i, " started.\n")

      if(DataFromSource$DR.File[i] == "NONE"){
        cat("\nNo DR file.")
      }else{


        if(DataFromSource$DR.File[i] == "OTHER"){
          cat("\nOther DR file.")
          file_name <- DataFromSource$DR.Other.file.name[i]
        }else{
          #Read in filename from Excel sheet:
          cat("\nDR file available.")

          file_name <-paste(DataFromSource[i,79], file_name, sep="_")
        }

        results <- extract_meas_data_from_xlsx(excel_file,
                                               excel_data_frame,
                                               meas_no = i,
                                               alpha_packet_size =alpha_packet_size)
        meancpa <- results$meancpa
        alpha_non_zero <- results$alpha_non_zero
        counts_pro_alpha <- results$counts_pro_alpha
        mfpga_bin <- results$mfpga_bin
        alph_start <- results$alph_start
        alph_end <- results$alph_end
        alpha_packet_size <- results$alpha_packet_size



        #-----------------------------------------------------------------------
        #Save analysis data of dark rate file into Excel file:
        #-----------------------------------------------------------------------
        #Number of ionizations per Trigger
        h_ion = graphics::hist(counts_pro_alpha, breaks= seq(0-0.5, max(counts_pro_alpha)+0.5, by=1))


        graphics::plot(1:length(counts_pro_alpha), counts_pro_alpha, xlab = "Alpha number", ylab = "Counts per trigger", main = "Number of counts per trigger")
        for (s in seq(from=(max(counts_pro_alpha)+2), to=33)){
          h_ion$counts[s] = 0
          #cat(s, ": ", h_ion$counts[s], "\n")
        }

        for (k in seq(from=1, to=33)){
          DataFromSource[i,87+k] <- as.numeric(h_ion$counts[k]) #No of k ionizations per Trigger
        }


        #---Filter for noise---------------------------------------
        filter_out_noise <- DataFromSource[i,127]
        if(is.na(filter_out_noise)){
          filter_out_noise <- 0
          DataFromSource[i,73] <- 0
          }
        #cat("Filter out noise: ", filter_out_noise)

        if(filter_out_noise == 1){
          #filter out noise
          filter_noise_above_x <- DataFromSource[i,128]

          mfpga_bin_df <- data.frame(mfpga_bin)

          #filter out rows where more ioniziations than filer_noise_above_x were recorded

          mfpga_bin_df_filtered <- mfpga_bin_df[mfpga_bin_df[,filter_noise_above_x]==0,]

          #convert df back to matrix
          mfpga_bin <- data.matrix(mfpga_bin_df_filtered)

          #Recalculate results based on mfpga_bin:
          counts_pro_alpha <- rowSums(mfpga_bin[alph_start:alph_end,] !=0)
          alpha_non_zero <- sum(rowSums(mfpga_bin[alph_start:alph_end,]) != 0)
          length_counts_pro_alpha = floor(length(counts_pro_alpha)/alpha_packet_size)*alpha_packet_size
          counts_pro_alpha <- counts_pro_alpha[1:length_counts_pro_alpha]
          mcpa <- matrix(counts_pro_alpha, ncol = alpha_packet_size, byrow = TRUE)
          meancpa <- rowSums(mcpa)/alpha_packet_size
          meancpa[-length(meancpa)] #remove last package because it might only be partially filled

        }

        #Write Mean Counts per Trigger over full measurement
        DataFromSource[i,81] <- as.numeric(mean(rowSums(mfpga_bin[,]!=0)))

        #Write Total No of alphas
        DataFromSource[i,82] <- as.numeric(nrow(mfpga_bin) )

        #Write Mean Counts per Trigger over selected measurement region
        DataFromSource[i,85] <- as.numeric(mean(counts_pro_alpha))

        #Write The percentage of Stable Signal
        DataFromSource[i,86] <- as.numeric(round((alph_end-alph_start)/nrow(mfpga_bin),3))



        #Calculate standard deviation of C/T spread
        DataFromSource$`DR.Standard.deviation.C/T.spread`[i] <- as.numeric(stats::sd(meancpa))


      }
      cat("\nDR Data Analysis written")
    }
    else if (DataFromSource$Data.analysis[i]==0 || is.na(DataFromSource$Data.analysis[i])==TRUE ){
      cat("\nDR Data Analysis skipped")
    }
  }
return(DataFromSource)
}
