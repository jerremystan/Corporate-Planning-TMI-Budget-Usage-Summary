library(dplyr)
library(readxl)
library(openxlsx)

df_template <- read_excel("Budget Usage 3Q 2024 - Template 1.xlsx", sheet = "JBD", range = "A2:K107")

df_source_a <- read_excel("WP Budget Report for HoD Q3 2024.xlsx", sheet = "Actual FY23", range = "A2:BH108")
df_source_b <- read_excel("WP Budget Report for HoD Q3 2024.xlsx", sheet = "BP24-Int", range = "A2:BH108")
df_source_c <- read_excel("WP Budget Report for HoD Q3 2024.xlsx", sheet = "Actual 23 Q3", range = "A2:BH108")
df_source_d <- read_excel("WP Budget Report for HoD Q3 2024.xlsx", sheet = "Actual 24 Q3", range = "A2:BH108")

dir.create("Output")

vector_dept <- colnames(df_source_a)[4:59]

vector_dept <- gsub("...50","", vector_dept)
vector_dept <- gsub("...20","", vector_dept)

#grep_GAD <- grep("GAD", vector_dept)

for (i in 1: length(vector_dept)){
  
  dept_name <- vector_dept[i]
  print(i)
  if (dept_name == "AB9" | dept_name == "AB2" | i == 47){
    
    #Do Nothing
    
  }
  
  else{
    
    if(i == 17){
      
      df_s_group <- cbind(df_source_a[,1:3],
                          df_source_a[,i+3]+df_source_a[,53]+df_source_a[,54],
                          df_source_b[,i+3]+df_source_b[,53]+df_source_b[,54],
                          df_source_c[,i+3]+df_source_c[,53]+df_source_c[,54],
                          df_source_d[,i+3]+df_source_d[,53]+df_source_d[,54])
      
    }
    else{
      
      df_s_group <- cbind(df_source_a[,1:3],df_source_a[,i+3],df_source_b[,i+3],df_source_c[,i+3],df_source_d[,i+3])
      
    }
    
    df_s_total <- cbind(df_source_a[,1:3],df_source_a$Total,df_source_b$Total,df_source_c$Total,df_source_d$Total)
    colnames(df_s_total)[2] <- "CoA Description"
    colnames(df_s_group)[2] <- "CoA Description"
    colnames(df_s_group)[4:7] <- c("Full Year Actual 2023", "Full Year BP 2024 - Internal", "2 Quarters Actual 2023", "2 Quarters Actual 2024")
    colnames(df_s_total)[4:7] <- c("Full Year Actual 2023", "Full Year BP 2024 - Internal", "2 Quarters Actual 2023", "2 Quarters Actual 2024")
  
    df_template_new <- df_template[,-4:-ncol(df_template)]
    
    dept_name <- c("Dept:",vector_dept[i])
    template_join <- left_join(df_template_new, df_s_group, by=colnames(df_template_new)[1:3])
    new_row <- c("TOTAL",NA,NA)
    for (j in 4:ncol(template_join)){
      
      sum_temp <- sum(template_join[,j])
      
      new_row <- c(new_row,sum_temp)
      
    }
    template_join <- rbind(template_join,new_row)
    template_join[,4:ncol(template_join)] <- lapply(template_join[, 4:ncol(template_join)], as.numeric)
    template_join$`...` <- NA
    template_join$`BP24 vs ACT23` <- with(template_join, 
                                          ifelse(
                                            is.finite((template_join$`Full Year BP 2024 - Internal`/template_join$`Full Year Actual 2023`) - 1),  # Check if the result is a valid number
                                            (template_join$`Full Year BP 2024 - Internal`/template_join$`Full Year Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                            ifelse(template_join$`Full Year BP 2024 - Internal`> template_join$`Full Year Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                          )
    )
    template_join$`1Q 24 vs 1Q 23` <- with(template_join, 
                                           ifelse(
                                             is.finite((template_join$`2 Quarters Actual 2024`/template_join$`2 Quarters Actual 2023`) - 1),  # Check if the result is a valid number
                                             (template_join$`2 Quarters Actual 2024`/template_join$`2 Quarters Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                             ifelse(template_join$`2 Quarters Actual 2024`> template_join$`2 Quarters Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                           )
    )
    template_join$`BP24 Usage` <- with(template_join, 
                                       ifelse(
                                         is.finite(template_join$`2 Quarters Actual 2024` / template_join$`Full Year BP 2024 - Internal`),  # Check if the division result is a valid number
                                         template_join$`2 Quarters Actual 2024` / template_join$`Full Year BP 2024 - Internal`,             # If valid, return the result
                                         NA                   # Otherwise, return an empty string
                                       )
    )
    
    
    sum_comp_name <- c("CT Code", "TMI")
    summary_comp <- df_s_total %>%
      group_by(`CoA Budget Group`) %>%
      summarise(
        `Full Year Actual 2023` <- sum(`Full Year Actual 2023`),
        `Full Year BP 2024 - Internal` <- sum(`Full Year BP 2024 - Internal`),
        `2 Quarters Actual 2023` <- sum(`2 Quarters Actual 2023`),
        `2 Quarters Actual 2024` <- sum(`2 Quarters Actual 2024`)
      )
    colnames(summary_comp)[2:5] <- c("Full Year Actual 2023", "Full Year BP 2024 - Internal", "2 Quarters Actual 2023", "2 Quarters Actual 2024")
    new_row <- c("TOTAL")
    for (j in 2:ncol(summary_comp)){
      
      sum_temp <- sum(summary_comp[,j])
      
      new_row <- c(new_row,sum_temp)
      
    }
    summary_comp <- rbind(summary_comp,new_row)
    summary_comp[,2:ncol(summary_comp)] <- lapply(summary_comp[, 2:ncol(summary_comp)], as.numeric)
    summary_comp$`...` <- NA
    summary_comp$`BP24 vs ACT23` <- with(summary_comp, 
                                          ifelse(
                                            is.finite((summary_comp$`Full Year BP 2024 - Internal`/summary_comp$`Full Year Actual 2023`) - 1),  # Check if the result is a valid number
                                            (summary_comp$`Full Year BP 2024 - Internal`/summary_comp$`Full Year Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                            ifelse(summary_comp$`Full Year BP 2024 - Internal`> summary_comp$`Full Year Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                          )
    )
    summary_comp$`1Q 24 vs 1Q 23` <- with(summary_comp, 
                                           ifelse(
                                             is.finite((summary_comp$`2 Quarters Actual 2024`/summary_comp$`2 Quarters Actual 2023`) - 1),  # Check if the result is a valid number
                                             (summary_comp$`2 Quarters Actual 2024`/summary_comp$`2 Quarters Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                             ifelse(summary_comp$`2 Quarters Actual 2024`> summary_comp$`2 Quarters Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                           )
    )
    summary_comp$`BP24 Usage` <- with(summary_comp, 
                                       ifelse(
                                         is.finite(summary_comp$`2 Quarters Actual 2024` / summary_comp$`Full Year BP 2024 - Internal`),  # Check if the division result is a valid number
                                         summary_comp$`2 Quarters Actual 2024` / summary_comp$`Full Year BP 2024 - Internal`,             # If valid, return the result
                                         NA                   # Otherwise, return an empty string
                                       )
    )
    
    
    sum_group_name <- c("CT Code", vector_dept[i])
    summary_group <- df_s_group %>%
      group_by(`CoA Budget Group`) %>%
      summarise(
        `Full Year Actual 2023` <- sum(`Full Year Actual 2023`),
        `Full Year BP 2024 - Internal` <- sum(`Full Year BP 2024 - Internal`),
        `2 Quarters Actual 2023` <- sum(`2 Quarters Actual 2023`),
        `2 Quarters Actual 2024` <- sum(`2 Quarters Actual 2024`)
      )
    colnames(summary_group)[2:5] <- c("Full Year Actual 2023", "Full Year BP 2024 - Internal", "2 Quarters Actual 2023", "2 Quarters Actual 2024")
    new_row <- c("TOTAL")
    for (j in 2:ncol(summary_group)){
      
      sum_temp <- sum(summary_group[,j])
      
      new_row <- c(new_row,sum_temp)
      
    }
    summary_group <- rbind(summary_group,new_row)
    summary_group[,2:ncol(summary_group)] <- lapply(summary_group[, 2:ncol(summary_group)], as.numeric)
    summary_group$`...` <- NA
    summary_group$`BP24 vs ACT23` <- with(summary_group, 
                                         ifelse(
                                           is.finite((summary_group$`Full Year BP 2024 - Internal`/summary_group$`Full Year Actual 2023`) - 1),  # Check if the result is a valid number
                                           (summary_group$`Full Year BP 2024 - Internal`/summary_group$`Full Year Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                           ifelse(summary_group$`Full Year BP 2024 - Internal`> summary_group$`Full Year Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                         )
    )
    summary_group$`1Q 24 vs 1Q 23` <- with(summary_group, 
                                          ifelse(
                                            is.finite((summary_group$`2 Quarters Actual 2024`/summary_group$`2 Quarters Actual 2023`) - 1),  # Check if the result is a valid number
                                            (summary_group$`2 Quarters Actual 2024`/summary_group$`2 Quarters Actual 2023`) - 1,             # If valid, return E3/D3 - 1
                                            ifelse(summary_group$`2 Quarters Actual 2024`> summary_group$`2 Quarters Actual 2023`, 1, NA)   # Otherwise, check if E3 > D3
                                          )
    )
    summary_group$`BP24 Usage` <- with(summary_group, 
                                      ifelse(
                                        is.finite(summary_group$`2 Quarters Actual 2024` / summary_group$`Full Year BP 2024 - Internal`),  # Check if the division result is a valid number
                                        summary_group$`2 Quarters Actual 2024` / summary_group$`Full Year BP 2024 - Internal`,             # If valid, return the result
                                        NA                   # Otherwise, return an empty string
                                      )
    )
    
    wb <- createWorkbook()
    
    addWorksheet(wb,"Summary")
    addWorksheet(wb,vector_dept[i])
    
    writeData(wb,"Summary",t(sum_comp_name),startCol = 1,startRow = 1,colNames = FALSE)
    writeData(wb,"Summary",summary_comp,startCol = 1,startRow = 2)
    
    writeData(wb,"Summary",t(sum_group_name),startCol = 1,startRow = 12,colNames = FALSE)
    writeData(wb,"Summary",summary_group,startCol = 1,startRow = 13)
    
    writeData(wb, vector_dept[i], t(dept_name),startCol = 1,startRow = 1,colNames = FALSE)
    writeData(wb,vector_dept[i], template_join,,startCol = 1,startRow = 2)
    
    saveWorkbook(wb, paste0("Output/Budget Usage 3Q 2024 - ",vector_dept[i],".xlsx"), overwrite = TRUE)
    
  
  }
}
