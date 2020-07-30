#' data_analysis_main
#' @description Extracts the main measurement data from a dataframe and then analyses it and saves the results
#' into the correct excel columns. Only executed if Data$Data.analysis[i]==1.
#'
#' @param excel_data_frame Dataframe of excel sheet to be analysed, obtained with \code{\link{extract_dataFrame_from_xlsx}}.
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#' @param alpha_packet_size Number of alpha particles to be grouped together into one data point,
#' default 100.
#' @importFrom("graphics", "abline", "hist", "plot")
#' @importFrom("dplyr", "filter")
#' @importFrom("stats", "sd")
#' @return Dataframe with main analysis completed.
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
data_analysis_main <- function(excel_file,
                               excel_data_frame,
                               alpha_packet_size=100


){
  DataFromSource <- excel_data_frame
  excel_nrows <- dim(excel_data_frame)[1]



  for( i in 1:(excel_nrows)){       # -1 because header is not included for filename writing
    cat("\n--- Measurement ", i, "-------------------------------------------------------------------")

    if(is.na(DataFromSource$Data.analysis[i])){
     DataFromSource$Data.analysis[i] <- 0}



    if (DataFromSource$Data.analysis[i]==1){ #Check if overwrite is activated

      #Extract measurement data of Measurement i from binary file
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


      #Mean Counts per Alpha Plot
      graphics::plot(1:length(meancpa), meancpa,
           xlab = paste("Consecutive means (alpha packet size = ",
                        alpha_packet_size, ")", sep = ""),
           ylab = "Mean counts per alpha",
           main = DataFromSource$Filename..txt[i],
           cex.main = 0.5)
      graphics::abline(h = mean(meancpa), col = "red")


      #Hist
      #Conversion of number entry in mfpga_bin into time:
      #1s = 50'000 -> number in mfpga_bin is divided through 50 and
      #multiplied with time window (time_window) to get time in ms

      #Read out time window from Excel file:
      time_window <- DataFromSource[i,8]

      # hist( mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]
      #       [ !mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]==0 ]/(50/time_window),
      #       xlab = "Arrival time [microseconds]",
      #       breaks = 25,
      #       main = DataFromSource$Filename..txt[i],
      #       cex.main = 0.5)



      #-----------------------------------------------------------------------
      #Save analysis data of main file into Excel file
      #-----------------------------------------------------------------------

      cat("\nData analysis of i=", i, " started.\n")

      #Number of ionizations per Trigger

      #Creates histogram with 0-32 ionizations. 1 bin per No of ionizations
      h_ion = graphics::hist(counts_pro_alpha, breaks= seq(0-0.5, max(counts_pro_alpha)+0.5, by=1))

      #Check for the highest number of ionizations per trigger, set everything above to 0.
      #Necessary, because otherwise the histogramm counts is NA not 0.
      for (s in seq(from=(max(counts_pro_alpha)+2), to=33)){
        h_ion$counts[s] = 0
      }

      #fill in the number of ionizations counted into excel file.
      for (k in seq(from=1, to=33)){
        DataFromSource[i,23+k] <- as.numeric(h_ion$counts[k]) #No of k ionizations per Trigger
      }


      #Write Total No of alphas
      DataFromSource[i,18] <- as.numeric(nrow(mfpga_bin) )

      #---Filter for noise
      filter_out_noise <- DataFromSource[i,73]
      if(is.na(filter_out_noise)){
        filter_out_noise <- 0
        DataFromSource[i,73] <- 0
        }
      #cat("Filter out noise: ", filter_out_noise)

      if(filter_out_noise == 1){
        #filter out noise
        filter_noise_above_x <- DataFromSource[i,74]

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
      DataFromSource[i,17] <- as.numeric(mean(rowSums(mfpga_bin[,]!=0)))

      #Write Mean Counts per Trigger over selected measurement region
      DataFromSource[i,21] <- as.numeric(mean(counts_pro_alpha))

      #Write The percentage of Stable Signal
      DataFromSource[i,22] <- as.numeric(round((alph_end-alph_start)/nrow(mfpga_bin),3))

      #Arrival time peak
      if(DataFromSource[i,21] != 0){
        breaks_hist <- 50*time_window

        h <- graphics::hist(x = mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]
                  [ !mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]==0 ]/(50/time_window),
                  breaks = breaks_hist)

        max_counts_hist = max(h$counts) #Find max counts
        arrival_time_peak <- 0
        for (m in seq(1:(breaks_hist))){ #Look for bin which contains max counts to get time
          if (h$counts[m]==max_counts_hist){
            arrival_time_peak <- h$mids[m]
            break
          }
        }}

      DataFromSource[i,23] <- as.numeric(arrival_time_peak)


      #Time between 2 ionizations

      #Test Matrix for overview
      # mfpga_bin <- matrix(data = c(0,0,0,0,0,
      #                              1,0,0,0,0,
      #                              2,4,0,0,0,
      #                              3,5,7,0,0,
      #                              4,5,6,7,0
      # ), ncol=5, byrow=TRUE)
      # print(mfpga_bin)


      df_mfpga_bin <- data.frame(mfpga_bin) #convert to data frame

      #Filter for at least 2 ionizations, reduces size of matrix by a lot.
      multiple_ion_mfpga_bin <- dplyr::filter(df_mfpga_bin, df_mfpga_bin$X2!=0)

      if(nrow(multiple_ion_mfpga_bin)==0){
        cat("Only single ionizations.")
      }else{

        if(nrow(multiple_ion_mfpga_bin)>10000){
          cat("\n Slow calculation due to many multiple ionizations (skipped): ",nrow(multiple_ion_mfpga_bin),"\n")
          #Only calculate first 10'000 ionization time differences
          #Initalize empty matrix to save time differences between ionizations later
          matrix_time_between_2_ionizations <- matrix(data =0 ,nrow=10000,
                                                      ncol=(ncol(multiple_ion_mfpga_bin)-1))

          #fill in matrix with time differences. Ex: Col 1 == Time between ionization 1 & 2
          for (s in seq(1:10000)){
            cat(".") #helps to see if code is still running :)
            for (k in seq(1:ncol(matrix_time_between_2_ionizations))){
              if (multiple_ion_mfpga_bin[s,k+1]!=0){ #if max of ionizations is reached do not calculate time difference
                matrix_time_between_2_ionizations[s,k] <- multiple_ion_mfpga_bin[s,k+1]-multiple_ion_mfpga_bin[s,k]
              }
            }
          }
          }else{

          #Initalize empty matrix to save time differences between ionizations later
          matrix_time_between_2_ionizations <- matrix(data =0 ,nrow=nrow(multiple_ion_mfpga_bin),
                                                      ncol=(ncol(multiple_ion_mfpga_bin)-1))

          #fill in matrix with time differences. Ex: Col 1 == Time between ionization 1 & 2
          for (s in seq(1:nrow(multiple_ion_mfpga_bin))){
            cat(".") #helps to see if code is still running :)
            for (k in seq(1:ncol(matrix_time_between_2_ionizations))){
              if (multiple_ion_mfpga_bin[s,k+1]!=0){ #if max of ionizations is reached do not calculate time difference
                matrix_time_between_2_ionizations[s,k] <- multiple_ion_mfpga_bin[s,k+1]-multiple_ion_mfpga_bin[s,k]
              }
            }
          }
          cat("\n")

          #Take the mean over all ionizations and ignore values which are equal to 0
          DataFromSource[i,57] <- mean(matrix_time_between_2_ionizations[matrix_time_between_2_ionizations!=0])/(50/time_window)

          DataFromSource[i,61] <- min(matrix_time_between_2_ionizations[matrix_time_between_2_ionizations!=0])/(50/time_window)

          #Mean over certain ionizations (1&2, 2&3,...) and ignore values=0.
          for (m in seq(from=1, to=3)){
            DataFromSource[i,57+m] <- mean(matrix_time_between_2_ionizations[,m]
                                           [matrix_time_between_2_ionizations[,m]!=0])/(50/time_window)

            if(identical(matrix_time_between_2_ionizations[,m],matrix(data=0, nrow=nrow(matrix_time_between_2_ionizations),
                                                                      ncol=ncol(matrix_time_between_2_ionizations))[,m])){
              DataFromSource[i,57+m] <- 0
            }
          }
        }
      }

      #Calculate standard deviation of C/T spread
      DataFromSource$`Standard.deviation.C/T.spread`[i] <- as.numeric(stats::sd(meancpa))

      #Alpha detection efficiency
      alpha_non_zero <- sum(rowSums(mfpga_bin[alph_start:alph_end,])!=0)
      DataFromSource[i,69] <- as.numeric(alpha_non_zero/nrow(mfpga_bin[alph_start:alph_end,]))

      #mean arrival time
      mean_arrival_time <- mean(mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]
                                [ !mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]==0 ]/(50/time_window))
      DataFromSource[i,70] <- as.numeric(mean_arrival_time)

      #sd arrival time
      sd_arrival_time <- stats::sd(mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]
                            [ !mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]==0 ]/(50/time_window))
      DataFromSource[i,71] <- as.numeric(sd_arrival_time)

      #sd arrival time / sqrt(N)
      DataFromSource[i,72] <- as.numeric(sd_arrival_time/sqrt(alph_end-alph_start))








      cat("\nData Analysis written")

      rm(multiple_ion_mfpga_bin,
         matrix_time_between_2_ionizations

         )
    }
    else if (DataFromSource$Data.analysis[i]==0){
      cat("\nData Analysis skipped")
    }
  }

  return (DataFromSource)
}
