# Written by Dr. Deepu Palal, free to use and distribute with citation

install_and_load <- function(pkg) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

# List of required packages
packages <- c("readxl", "stats", "writexl", "officer", "flextable", 
              "dplyr", "nortest", "DescTools")

sapply(packages, install_and_load)
options(nwarnings = 10000)

#output_location<-c("C:/Users/user/Downloads/Descriptives1.docx")
doc <- read_docx()
formatted_text <- fpar(ftext("Analysis", 
                             prop = fp_text(font.size = 24, 
                                            font.family = "Times New Roman", 
                                            bold = TRUE)))
doc <- body_add_fpar(doc, formatted_text, style = "centered")

doc <- body_add_par(doc, paste0("Date and Time: ",Sys.time()), style = "centered")
doc <- body_add_par(doc, " ", style = "centered")
tableno<-1

## Functions ####
calculate_frequency <- function(freqcolumns) {
  for (colno_colno_cat_freq_columns in freqcolumns)
  {
    results_cat_freq_columns <- data.frame(
      Category = character(0),
      Frequency = integer(0),
      Percentage = numeric(0),
      CI_95 = character(0)
    )
    frequency <- table(df1[[colno_colno_cat_freq_columns]])
    percent <- prop.table(frequency) * 100
    n <- sum(frequency)
    for (subcategory_colno_colno_cat_freq_columns in unique(df1[[colno_colno_cat_freq_columns]][!is.na(df1[[colno_colno_cat_freq_columns]]) & is.character(df1[[colno_colno_cat_freq_columns]])])) {
      category_frequency <- frequency[subcategory_colno_colno_cat_freq_columns]
      ci_95 <- binom.test(category_frequency, n, conf.level = 0.95)$conf.int
      results_cat_freq_columns <- rbind(
        results_cat_freq_columns,
        data.frame(
          Category = subcategory_colno_colno_cat_freq_columns,
          Frequency = category_frequency,
          Percentage = round(percent[subcategory_colno_colno_cat_freq_columns],2),
          CI_95 = paste0(round(ci_95[1]*100,2)," - ",round(ci_95[2]*100,2))
        )
      )
    }
    colnames(results_cat_freq_columns) <- c(names(df1)[[colno_colno_cat_freq_columns]],"Frequency","Percentage","CI_95")
    total_row <- data.frame(
      Category = "Total",
      Frequency = n,
      Percentage = 100,
      CI_95 = ""
    )
    colnames(total_row) <- c(names(df1)[[colno_colno_cat_freq_columns]],"Frequency","Percentage","CI_95")
    results_cat_freq_columns <- rbind(results_cat_freq_columns, total_row)
    
    tablename <- paste0("Table ", tableno, ": ", c(names(df1)[[colno_colno_cat_freq_columns]]))
    doc <- body_add_par(doc, tablename)
    doc <- body_add_par(doc, "")
    std_border = fp_border(color="gray")
    desc_table <- flextable(results_cat_freq_columns) 
    desc_table <- desc_table  %>% autofit() %>% 
      fit_to_width(20, inc = 1L, max_iter = 2, unit = "cm")  %>%
      border_remove() %>% 
      theme_booktabs() %>%
      vline(part = "all", j = 3, border = NULL) %>%
      flextable::align(align = "center", j = c(3:4), part = "all") %>% 
      flextable::align(align = "center", j = c(1:2), part = "header") %>%
      fontsize(i = 1, size = 10, part = "header") %>%   # adjust font size of header
      bold(i = 1, bold = TRUE, part = "header") %>%     # adjust bold face of header
      bold(j = 1, bold = TRUE, part = "body")  %>% 
      merge_v(j = 1)%>%
      merge_v(j = 2)%>%
      hline(part = "body", border = std_border)
    doc <- body_add_flextable(doc, value = desc_table)
    doc <- body_add_par(doc, "")
    tableno<<-tableno+1
  }
  rm(list= c("category_frequency","ci_95","colno_colno_cat_freq_columns","desc_table",                              
             "freqcolumns","frequency","percent","n", "results_cat_freq_columns","std_border","subcategory_colno_colno_cat_freq_columns","tablename",                               
             "total_row","typeoffacility"))
}
association <- function(dependent, independent, tablepercent){
  if (is.null(dependent)) { } else {
    for (independentvar in independent) {
      independent_name <- names(df1)[[independentvar]]
      #doc <- body_add_par(doc, independent_name, style = "heading 1")
      for (dependentvar in dependent)
      {
        dependent_name <- names(df1)[[dependentvar]]
        tabledesc <- paste0("Table ", tableno, ": Association between ", independent_name, " and ", dependent_name)
        doc <- body_add_par(doc, "", style = "heading 2")
        
        cont_table <- table(df1[[independentvar]], df1[[dependentvar]])
        expected_counts <- chisq.test(cont_table)$expected
        prop_low_counts <- sum(expected_counts < 5) / length(expected_counts)
        extranote<-c("")
        if (prop_low_counts > 0.2 || any(chisq.test(cont_table)$expected < 1)) {
          fisher_test <- tryCatch(fisher.test(cont_table), error = function(e) NULL)
          if (is.null(fisher_test)) {
            fisher_test <- fisher.test(cont_table, simulate.p.value = TRUE, B=1e4)
            extranote <- c("Using Monte Carlo simulation")
          }
          pvalue<-round(fisher_test["p.value"][[1]],3)
          finalp<- paste0("Fisher exact P=",replace(pvalue,pvalue==0,"<0.001")," ",extranote)
        } else {
          chisq_test <- chisq.test(cont_table)  
          chisq<-chisq_test["statistic"][[1]][[1]]
          dfreedom<-chisq_test["parameter"][[1]][[1]]
          pvalue<-round(chisq_test["p.value"][[1]],3)
          finalp<- paste0("chisq(",dfreedom,")=",round(chisq,3),", P=",replace(pvalue,pvalue==0,"<0.001"))
        }
        
        cont_table_old <- cbind(cont_table, Total = rowSums(cont_table))
        cont_table_old <- rbind(cont_table_old, Total = colSums(cont_table_old))
        
        if (tablepercent=="row") {
          prop_table <- proportions(cont_table, margin = 1) * 100 # row percentages
        } else {
          if (tablepercent=="column") {
            prop_table <- proportions(cont_table, margin = 2) * 100 # row percentages
          }
        }
        
        col_pct <- rowSums(cont_table) / sum(cont_table) * 100  
        row_pct <- colSums(cont_table) / sum(cont_table) * 100
        prop_table <- rbind(prop_table, Total = row_pct)  
        prop_table <- cbind(prop_table, Total = col_pct)  
        prop_table[nrow(prop_table), ncol(prop_table)] <- 100
        
        
        table_print <- matrix(paste0(cont_table_old, " (", round(prop_table,2),")"), nrow = nrow(prop_table))
        rownames(table_print) <- rownames(cont_table_old)
        colnames(table_print) <- colnames(cont_table_old)
        
        
        table_print <- as.data.frame.matrix(table_print)
        table_print$Categories <- rownames(table_print)
        table_print <- table_print[, c(ncol(table_print), 1:(ncol(table_print)-1))]
        
        ##Printing Chisquare table using Flextable####
        nobrackets <- rep("N (%)", times = ncol(table_print)-1 )
        nobracketsfinal <- c("",nobrackets)
        
        chi_table <- flextable(table_print)
        chi_table <- chi_table  %>% autofit() %>% 
          fit_to_width(20, inc = 1L, max_iter = 20, unit = "cm")  %>% 
          add_header_row(top = FALSE,
                         values = nobracketsfinal) %>%
          border_remove() %>% 
          theme_booktabs() %>%
          vline(part = "all", j = 1, border = NULL) %>%   # at column 2 
          vline(part = "all", j = ncol(table_print)-1, border = NULL)  %>%
          hline(i = nrow(table_print)-1, border = NULL, part = "body")  %>%
          add_footer_row(values = c(finalp),
                         colwidths = ncol(table_print), top = TRUE) %>%
          fontsize(i = 1, size = 10, part = "footer") %>%
          hline(i = 1, part = "header", border = fp_border(color="transparent"))  %>%
          set_caption(tabledesc)
        chi_table
        
        doc <- body_add_flextable(doc, value = chi_table)
        #doc <- body_add_par(doc, finalp, style = "Normal")
        rm(list= c("cont_table", "expected_counts", "prop_low_counts", "finalp", "chisq_test", "chisq", "dfreedom", "pvalue", "cont_table_old", "col_pct", "row_pct", "prop_table", "table_print", "tabledesc", "chi_table","nobrackets", "nobracketsfinal"))
        #rm(list= c("cont_table", "expected_counts", "prop_low_counts", "fisher_test", "pvalue", "finalp", "chisq_test", "chisq", "dfreedom", "pvalue", "cont_table_old", "prop_table", "col_pct", "row_pct", "prop_table", "table_print", "tabledesc", "chi_table","nobrackets", "nobracketsfinal"))
        tableno<<-tableno+1
      }
    }
  }
}

##Functions Commands####

#calculate_frequency (freqcolumns<- c())
#association (dependent<-c(), independent<-c(), tablepercent<-c("row"))

## Print Table ####
#print(doc, target = output_location)
