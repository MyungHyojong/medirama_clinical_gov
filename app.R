# Load required packages
library(shiny)
library(httr)
library(jsonlite)
library(readr)
library(dplyr)
library(openxlsx)
library(stringr)
library(purrr)
library(zip)
library(knitr) 
library(kableExtra)
library(gemini.R)
library(DT)
library(officer)    # For generating PowerPoint presentations
library(flextable)  # For formatting tables in PowerPoint
library(htmltools)
library(openai)  # OpenAI API 패키지 추가

openai_api_key <- "YOUR API KEY"
Sys.setenv(OPENAI_API_KEY = openai_api_key)

# Define helper functions
find_common_substring <- function(s1, s2) {
  if (is.null(s1) || is.null(s2) || nchar(s1) == 0 || nchar(s2) == 0) {
    return("")
  }
  max_len <- min(nchar(s1), nchar(s2))
  longest_common_substr <- ""
  for (len in seq(max_len, 1, -1)) {
    start_indices_s1 <- 1:max(1, nchar(s1) - len + 1)
    end_indices_s1 <- len:nchar(s1)
    substrs_s1 <- str_sub(s1, start_indices_s1, end_indices_s1)
    
    start_indices_s2 <- 1:max(1, nchar(s2) - len + 1)
    end_indices_s2 <- len:nchar(s2)
    substrs_s2 <- str_sub(s2, start_indices_s2, end_indices_s2)
    
    common_substrs <- intersect(substrs_s1, substrs_s2)
    if (length(common_substrs) > 0) {
      longest_common_substr <- common_substrs[1]
      break
    }
  }
  return(longest_common_substr)
}

extract_last_in_brackets <- function(text) {
  matches <- regmatches(text, gregexpr("\\(([^)]+)\\)", text))
  if (length(matches[[1]]) > 0) {
    last_match <- tail(matches[[1]], 1)
    return(last_match)
  } else {
    return(NA)
  }
}

# Define UI
ui <- navbarPage(
  title = "Clinical Trial Data Processor",
  tabPanel(
    "Upload and Process",
    sidebarLayout(
      sidebarPanel(
        fileInput("csv_file", "Upload CSV File", accept = c(".csv")),
        actionButton("process", "Process Data"),
        br(),
        downloadButton("download_df", "Download Updated Data"),
        br(),
        downloadButton("download_html", "Download HTML Files")
      ),
      mainPanel(
        DTOutput("df_table"),
        br(),
        h3("GPT Response"),
        verbatimTextOutput("gemini_response")
      )
    )
  ),
  tabPanel(
    "Select Rows and Generate PPT",
    sidebarLayout(
      sidebarPanel(
        fileInput("ppt_csv_file", "Upload CSV File", accept = c(".csv")),
        textInput("common_title", "Common Title for All Slides", value = "", placeholder = "Enter a common title for all slides"),
        uiOutput("grouping_column_selector"),
        sliderInput("rows_per_slide", "Number of Rows per Slide", min = 1, max = 20, value = 6, step = 1),
        uiOutput("column_selector"),
        uiOutput("column_width_sliders"),
        actionButton("generate_ppt", "Generate PPT"),
        downloadButton("download_ppt", "Download PPT")
      ),
      mainPanel(
        DTOutput("filtered_table"),
        br(),
        h3("Previews of Dataframes for Each Category"),
        uiOutput("slide_previews")
      )
    )
  )
)

# Define Server
server <- function(input, output, session) {
  rv <- reactiveValues(
    df = NULL,
    html_contents = list(),
    html_file_names = NULL,
    gemini_responses = list(),
    filtered_df = NULL,
    selected_rows = NULL,
    ppt_file = NULL,
    preview_data_list = list(),
    flextable_preview = NULL,
    df_ppt = NULL
  )
  
  # Existing code for the first tab remains unchanged
  observeEvent(input$process, {
    req(input$csv_file)
    # Add progress indicator
    withProgress(message = "Processing Data...", value = 0, {
      df <- read_csv(input$csv_file$datapath, locale = locale(encoding = 'utf-8'))
      #df <- df[201:nrow(df),]
      # Check if 'NCT Number' column exists
      if (!"NCT Number" %in% colnames(df)) {
        showNotification("Error: 'NCT Number' column not found in the uploaded CSV file.", type = "error")
        return(NULL)
      }
      
      nct_list <- df$`NCT Number`
      
      # Initialize result vectors
      tests <- c()
      reasons <- c()
      explanations <- c()
      genes <- c()
      confidence_scores <- c()
      experimental_1s <- c()
      experimental_2s <- c()
      control_1s <- c()
      control_2s <- c()
      study_names <- c()
      start_dates <- c()
      primary_completion_dates <- c()
      completion_dates <- c()
      html_contents <- list()
      html_file_names <- c()
      gemini_responses <- list()
      
      total_ncts <- length(nct_list)
      progress_step <- 1 / max(total_ncts, 1)
      current_progress <- 0
      
      # Process each NCT ID
      for (nct_id in nct_list) {
        # Wrap the processing code in tryCatch to handle errors
        print(nct_id)
        tryCatch({
          # API endpoint
          url <- paste0("https://clinicaltrials.gov/api/v2/studies/", nct_id)
          # GET request
          response <- GET(url, query = list(
            format = 'json',
            markupFormat = 'markdown',
            fields = 'NCTId,BriefTitle,ConditionsModule,EligibilityModule,OfficialTitle,ArmsInterventionsModule,StartDate,PrimaryCompletionDate,CompletionDate'
          ))
          if (status_code(response) == 200) {
            study_data <- content(response, as = "parsed")
          } else {
            print(paste("Error:", status_code(response), content(response, as = "text")))
            # Assign NA to variables and continue
            tests <- c(tests, NA)
            reasons <- c(reasons, NA)
            explanations <- c(explanations, NA)
            genes <- c(genes, NA)
            confidence_scores <- c(confidence_scores, NA)
            experimental_1s <- c(experimental_1s, NA)
            experimental_2s <- c(experimental_2s, NA)
            control_1s <- c(control_1s, NA)
            control_2s <- c(control_2s, NA)
            study_names <- c(study_names, NA)
            html_contents[[nct_id]] <- NA
            gemini_responses[[nct_id]] <- NA
            next
          }
          # Extract armGroups
          arm_groups <- study_data$protocolSection$armsInterventionsModule$armGroups
          if (is.null(arm_groups) || length(arm_groups) == 0) {
            # Assign NA to variables and continue
            #tests <- c(tests, NA)
            #reasons <- c(reasons, NA)
            #explanations <- c(explanations, NA)
            #genes <- c(genes, NA)
            #confidence_scores <- c(confidence_scores, NA)
            experimental_1s <- c(experimental_1s, NA)
            experimental_2s <- c(experimental_2s, NA)
            control_1s <- c(control_1s, NA)
            control_2s <- c(control_2s, NA)
            #study_names <- c(study_names, NA)
            #html_contents[[nct_id]] <- NA
            #gemini_responses[[nct_id]] <- NA
          } else{
            arms_data <- bind_rows(lapply(arm_groups, function(arm) {
              df_arm <- as.data.frame(arm, stringsAsFactors = FALSE)
              
              # Ensure 'label' exists
              if (!('label' %in% names(df_arm))) {
                df_arm$label <- NA
              }
              
              # Ensure 'type' exists
              if (!('type' %in% names(df_arm))) {
                df_arm$type <- NA
              }
              
              # Ensure 'description' exists
              if (!('description' %in% names(df_arm))) {
                df_arm$description <- NA
              }
              
              df_arm <- df_arm %>% select(label, type, description, everything())
              
              if (ncol(df_arm) > 4) {
                df_arm$col4 <- apply(df_arm[, 4:ncol(df_arm)], 1, function(x) paste(x, collapse = "/"))
                df_arm <- df_arm[, c(1:3, ncol(df_arm))]
              }
              
              if (ncol(df_arm) >= 4) {
                names(df_arm)[4] <- 'drug'
              }
              
              return(df_arm)
            }))
            
            # Initialize lists to store experimental and control data
            experimental_labels_drugs <- c()
            experimental_descriptions <- c()
            control_labels_drugs <- c()
            control_descriptions <- c()
            
            # Iterate over arms_data
            for (i in 1:nrow(arms_data)) {
              # Extract label, drug, description
              label <- arms_data$label[i]
              drug <- ifelse(is.null(arms_data$drug[i]), "No drug info", arms_data$drug[i]) # handle missing drug info
              description <- arms_data$description[i]
              
              # Classify as experimental or control based on type
              if (grepl('EXPERIMENTAL', arms_data$type[i])) {
                # Experimental
                labels_drugs <- paste(label, drug, sep = "\n")
                experimental_labels_drugs <- c(experimental_labels_drugs, labels_drugs)
                experimental_descriptions <- c(experimental_descriptions, description)
              } else {
                # Control
                labels_drugs <- paste(label, drug, sep = "\n")
                control_labels_drugs <- c(control_labels_drugs, labels_drugs)
                control_descriptions <- c(control_descriptions, description)
              }
            }
            
            # Combine experimental and control data
            experimental_data_labels_drugs <- paste(experimental_labels_drugs, collapse = "\n\n")
            experimental_data_descriptions <- paste(experimental_descriptions, collapse = "\n\n")
            control_data_labels_drugs <- paste(control_labels_drugs, collapse = "\n\n")
            control_data_descriptions <- paste(control_descriptions, collapse = "\n\n")
            experimental_1s <- c(experimental_1s, experimental_data_labels_drugs)
            experimental_2s <- c(experimental_2s, experimental_data_descriptions)
            control_1s <- c(control_1s, control_data_labels_drugs)
            control_2s <- c(control_2s, control_data_descriptions)
            
          }
          
          
          # Extract experiment description
          experiment_description <- toString(study_data)
          question <- paste0(
            'Please provide an answer in the following format based on the provided experiment description,\n',
            '1. test: Determine which of the following categories the cancer treatment experiment belongs to: first line, second line, third line, neoadjuvant, adjuvant or unclear. Only print out the test type without writinig any sentences. First line is the initial standard treatment, second line is the alternative treatment used after the failure of the first, third line is given after both first and second lines fail, neoadjuvant is treatment given before surgery to shrink the tumor, and adjuvant is given after surgery to prevent recurrence.\n',
            '2. reason: write the exact specific part of the eligibilityCriteria in Experiment description below that supports your answer in 1 "without changing a single part".\n',
            '3. explanations: explain specifically why you chose your answer in 1\n',
            '4. genes: mutations, expressions associated in study such as KRAS, EGFR, MET, ALK, CEACAM5, STK11, KEAP1\n',
            '5. Confidence Score: Based on the probability of incorrectly guessing the type of test written in response number 1, if you are uncertain about your decision, to determine test 1 answer write "uncertain". else write "certain".\n\n',
            '\nExperiment description is written below\n', 
            experiment_description
          )
          
          # OpenAI ChatGPT API를 사용하여 질문에 대한 응답 받기
          completion <- openai::create_chat_completion(
            model = "gpt-3.5-turbo",
            messages = list(
              list(role = "system", content = "You are a helpful assistant that analyzes clinical trial descriptions."),
              list(role = "user", content = question)
            )
          )
          response_text <- completion$choices$message.content[1]
          
          # Parse the response
          lines <- str_split(response_text, "\\n")[[1]]
          
          # Initialize variables
          test <- NA
          reason <- NA
          explanation <- NA
          genes_data <- NA
          confidence_score <- NA
          
          # Extract information from response
          for (line in lines) {
            line <- str_trim(line)  # Trim whitespace from the line
            if (str_detect(line, "^1\\.")) {
              test <- str_remove(line, "^1\\.\\s*")
              test <- gsub('test:','',test)
              test <- tolower(test)
            } else if (str_detect(line, "^2\\.")) {
              reason <- str_remove(line, "^2\\.\\s*")
            } else if (str_detect(line, "^3\\.")) {
              explanation <- str_remove(line, "^3\\.\\s*")
            } else if (str_detect(line, "^4\\.")) {
              genes_data <- str_remove(line, "^4\\.\\s*")
            } else if (str_detect(line, "^5\\.")) {
              confidence_score <- str_remove(line, "^5\\.\\s*")
            }
          }
          
          #print(test)
          tests <- c(tests, test)
          reasons <- c(reasons, reason)
          explanations <- c(explanations, explanation)
          genes <- c(genes, genes_data)
          confidence_scores <- c(confidence_scores, confidence_score)
          
          # Data processing and highlighting
          official_title <- study_data$protocolSection$identificationModule$officialTitle
          if(is.null(official_title)){ official_title <- 'NA'}
          brief_title <- study_data$protocolSection$identificationModule$briefTitle
          conditions <- paste(study_data$protocolSection$conditionsModule$conditions, collapse = ', ')
          eligibility_criteria <- study_data$protocolSection$eligibilityModule$eligibilityCriteria
          
          study_name <- ifelse(is.na(official_title),'na',extract_last_in_brackets(official_title))
          
          print(official_title)
          print(study_name)
          print('---------------------')
          #start_date <- 'not available'
          start_date <- unlist(study_data$protocolSection$statusModule$startDateStruct)
          primary_completion_date <- study_data$protocolSection$statusModule$primaryCompletionDateStruct
          completion_date <- study_data$protocolSection$statusModule$completionDateStruct

          if(is.null(start_date)){start_date <- 'NA'}
          if(is.na(start_date)){start_date <- 'NA'}
          if(is.null(completion_date)){completion_date <- 'NA'}
          if(is.na(completion_date)){completion_date <- 'NA'}
          if(is.null(primary_completion_date)){primary_completion_date <- 'NA'}
          if(is.na(primary_completion_date)){primary_completion_date <- 'NA'}
          
                    
          study_names <- c(study_names, study_name)
          start_dates <- c(start_dates, start_date)
          
          primary_completion_dates <- c(primary_completion_dates, primary_completion_date)
          completion_dates <- c(completion_dates, completion_date)
          reason_parts <- find_common_substring(reason, eligibility_criteria)
          # Highlighting
          highlighted_eligibility_criteria <- eligibility_criteria
          if (nchar(reason_parts) > 0) {
            highlighted_eligibility_criteria <- str_replace_all(
              highlighted_eligibility_criteria, 
              fixed(reason_parts), 
              paste0("<mark>", reason_parts, "</mark>")
            )
          }
          highlighted_eligibility_criteria <- gsub("\n", "<br>", highlighted_eligibility_criteria)
          plan_table <- kable(arms_data, format = "html", table.attr = "class='dataframe'") %>%
            kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "bordered"))
          
          # Markdown content
          markdown_content <- paste0(
            "<h1>Response</h1>\n\n",
            "<strong>Test:</strong> ", test, "<br><br>\n",
            "<strong>Explanation:</strong> ", explanation, "<br><br>\n",
            "<strong>Genes:</strong> ", genes_data, "<br><br>\n",
            "<strong>Confidence Score:</strong> ", confidence_score, "<br><br>\n",
            "<h1>Study Plan</h1>\n\n",
            plan_table,
            "\n<h1>Study Data</h1>\n\n",
            "<strong>Official Title:</strong> ", official_title, "<br><br>\n",
            "<strong>Brief Title:</strong> ", brief_title, "<br><br>\n",
            "<strong>Conditions:</strong> ", conditions, "<br><br>\n",
            "<strong>Eligibility Module:</strong><br>", highlighted_eligibility_criteria, "<br><br>\n",
            "<strong>Highlighted Reason from Response:</strong><br>", ifelse(nchar(reason_parts) > 0, paste("<mark>", reason_parts, "</mark>", collapse = "<br>"), "N/A")
          )
          # Store the HTML content and file names
          html_contents[[nct_id]] <- markdown_content
          html_file_names <- c(html_file_names, paste0(nct_id, ".html"))
          gemini_responses[[nct_id]] <- response_text
        }, error = function(e) {
          # Handle errors by assigning NA and moving on
          tests <- c(tests, NA)
          reasons <- c(reasons, NA)
          explanations <- c(explanations, NA)
          genes <- c(genes, NA)
          confidence_scores <- c(confidence_scores, NA)
          experimental_1s <- c(experimental_1s, NA)
          experimental_2s <- c(experimental_2s, NA)
          control_1s <- c(control_1s, NA)
          control_2s <- c(control_2s, NA)
          study_names <- c(study_names, NA)
          start_dates <- c(start_dates, NA)
          primary_completion_dates <- c(primary_completion_dates, NA)
          completion_dates <- c(completion_dates, NA)
          html_contents[[nct_id]] <- NA
          gemini_responses[[nct_id]] <- NA
          # Print the error message for debugging (optional)
          print(paste("Error processing NCT ID", nct_id, ":", e$message))
          # Continue to the next iteration
          next
        })
        # Update progress
        current_progress <- current_progress + progress_step
        incProgress(progress_step)
      }
      print(study_names)
      # Update df with new columns
      df <- df %>% 
        mutate(
          TestLines = tests,
          Reason = reasons,
          Explanation = explanations,
          Genes = genes,
          Confidence_Score = confidence_scores,
          experimental_data = experimental_1s,
          experimental_description= experimental_2s,
          control_datas = control_1s,
          control_description = control_2s, 
          study_name = study_names,
          start_date = start_dates,
          primary_completion_date = primary_completion_dates,
          completion_date = completion_dates
        )
      # Store the updated df and html contents in reactive values
      rv$df <- df
      rv$html_contents <- html_contents
      rv$html_file_names <- html_file_names
      rv$gemini_responses <- gemini_responses
      
      # Notify the user
      showNotification("Data processing complete.", type = "message")
    })
  })
  
  # Truncate long text to 20 characters
  truncate_text <- function(text, max_length = 20) {
    ifelse(nchar(text) > max_length, 
           paste0(substr(text, 1, max_length), "..."), 
           text)
  }
  
  # Output the data table with truncated text
  output$df_table <- renderDT({
    req(rv$df)
    df_display <- rv$df
    
    # Apply truncation to all character columns
    char_cols <- sapply(df_display, is.character)
    df_display[, char_cols] <- lapply(df_display[, char_cols], truncate_text)
    
    datatable(df_display, options = list(scrollX = TRUE))
  })
  
  # Download handler for df
  output$download_df <- downloadHandler(
    filename = function() {
      "updated_ctg_studies.csv"
    },
    content = function(file) {
      # list 열을 문자열로 변환
      clean_df <- rv$df %>%
        mutate(across(where(is.list), ~ sapply(.x, function(item) {
          if (is.null(item)) {
            return(NA)
          } else {
            paste(item, collapse = ", ")
          }
        })))
      
      # CSV로 저장
      write.csv(clean_df, file = file, row.names = FALSE)
    }
  )
  
  # Download handler for html files
  output$download_html <- downloadHandler(
    filename = function() {
      "html_files.zip"
    },
    content = function(file) {
      temp_dir <- tempdir()
      html_paths <- c()
      for (nct_id in names(rv$html_contents)) {
        html_content <- rv$html_contents[[nct_id]]
        if (is.na(html_content)) {
          next  # Skip if html_content is NA
        }
        html_file <- file.path(temp_dir, paste0(nct_id, ".html"))
        writeLines(html_content, con = html_file)
        html_paths <- c(html_paths, html_file)
      }
      if (length(html_paths) == 0) {
        # If no HTML files were generated, create an empty zip
        file.create(file)
      } else {
        zipr(zipfile = file, files = html_paths)
      }
    },
    contentType = "application/zip"
  )
  
  # Output the gemini responses
  output$gemini_response <- renderText({
    req(rv$gemini_responses)
    responses <- unlist(rv$gemini_responses)
    responses <- responses[!is.na(responses)]  # Remove NA values
    if (length(responses) == 0) {
      return("No responses available.")
    } else {
      paste(responses, collapse = "\n\n")
    }
  })
  
  # New code for the second tab starts here
  observeEvent(input$ppt_csv_file, {
    req(input$ppt_csv_file)
    df_ppt <- read_csv(input$ppt_csv_file$datapath, locale = locale(encoding = 'utf-8'))
    rv$df_ppt <- df_ppt
  })
  
  # Output the data table with truncated text for the second tab
  output$filtered_table <- renderDT({
    req(rv$df_ppt)
    df_ppt_display <- rv$df_ppt
    
    # Apply truncation to all character columns
    char_cols <- sapply(df_ppt_display, is.character)
    df_ppt_display[, char_cols] <- lapply(df_ppt_display[, char_cols], truncate_text)
    
    datatable(df_ppt_display, options = list(scrollX = TRUE))
  })
  
  # Generate column selector UI
  output$column_selector <- renderUI({
    req(rv$df_ppt)
    column_names <- names(rv$df_ppt)
    checkboxGroupInput("selected_columns", "Select Columns", choices = column_names, selected = column_names[1:min(3, length(column_names))])
  })
  
  # Generate grouping column selector UI
  output$grouping_column_selector <- renderUI({
    req(rv$df_ppt)
    column_names <- names(rv$df_ppt)
    selectInput("grouping_column", "Select Grouping Column (Optional)", choices = c("None", column_names), selected = "None")
  })
  
  # Generate individual sliders for each selected column
  output$column_width_sliders <- renderUI({
    req(input$selected_columns)
    
    # Calculate initial percentages equally divided
    num_cols <- length(input$selected_columns)
    initial_percentage <- floor(100 / num_cols)
    remaining <- 100 - initial_percentage * num_cols
    initial_values <- rep(initial_percentage, num_cols)
    if (remaining > 0) {
      initial_values[1:remaining] <- initial_values[1:remaining] + 1
    }
    
    sliders <- lapply(seq_along(input$selected_columns), function(idx) {
      col <- input$selected_columns[idx]
      sliderInput(
        inputId = paste0("width_", make.names(col)), 
        label = paste("Width (%) for", col), 
        min = 0, max = 100, value = initial_values[idx], step = 1
      )
    })
    do.call(tagList, sliders)
  })
  
  observeEvent(input$generate_ppt, {
    req(rv$df_ppt)
    withProgress(message = "Generating PPT...", value = 0, {
      common_title <- input$common_title  # 사용자가 입력한 공통 제목 가져오기
      
      if (input$grouping_column != "None") {
        grouping_column <- input$grouping_column
        categories <- unique(rv$df_ppt[[grouping_column]])
        df_list <- lapply(categories, function(cat) {
          rv$df_ppt %>% filter(!!sym(grouping_column) == cat) %>% select(all_of(input$selected_columns))
        })
        names(df_list) <- categories
      } else {
        df_list <- list(rv$df_ppt %>% select(all_of(input$selected_columns)))
        names(df_list) <- "Data"
      }
      
      # Collect column percentages
      col_percentages <- sapply(input$selected_columns, function(col) {
        input[[paste0("width_", make.names(col))]]
      })
      names(col_percentages) <- input$selected_columns
      
      total_percentage <- sum(col_percentages)
      if (total_percentage == 0) {
        showNotification("Total column width percentages cannot be zero.", type = "error")
        return(NULL)
      }
      col_percentages <- col_percentages / total_percentage * 100
      
      total_width <- 10
      col_widths <- col_percentages / 100 * total_width
      names(col_widths) <- input$selected_columns
      
      ppt <- read_pptx()
      total_steps <- sum(sapply(df_list, function(df) ceiling(nrow(df)/input$rows_per_slide)))
      progress_step <- 1 / max(total_steps, 1)
      current_progress <- 0
      
      rv$preview_data_list <- list()
      
      for (i in seq_along(df_list)) {
        df <- df_list[[i]]
        if (nrow(df) == 0) next
        
        num_rows_per_slide <- input$rows_per_slide
        num_splits <- ceiling(nrow(df) / num_rows_per_slide)
        df_splits <- split(df, rep(1:num_splits, each = num_rows_per_slide, length.out = nrow(df)))
        
        for (j in seq_along(df_splits)) {
          df_chunk <- df_splits[[j]]
          slide_title <- paste(common_title, names(df_list)[i], if (num_splits > 1) paste(" - page", j) else "")
          
          ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
          ppt <- ph_with(
            ppt, 
            value = fpar(ftext(slide_title, fp_text(font.size = 16, bold = TRUE))), 
            location = ph_location_type(type = "title")
          )
          
          flextable_df <- flextable(df_chunk) %>%
            border(border = fp_border(color = "black", width = 1), part = "all") %>%
            bg(part = "header", bg = "#d9f2d0")
          
          for (col in input$selected_columns) {
            col_width <- col_widths[col]
            flextable_df <- width(flextable_df, j = col, width = as.numeric(col_width))
          }
          
          flextable_df <- flextable_df %>%
            align(align = "left", part = "all") %>%
            height_all(height = 0.5)
          
          ppt <- ph_with(ppt, value = flextable_df, location = ph_location(left = 0.5, top = 1.5, width = 9, height = 5))
          
          if (j == 1) {
            rv$preview_data_list[[names(df_list)[i]]] <- df_chunk
          }
          
          current_progress <- current_progress + progress_step
          incProgress(progress_step)
        }
      }
      
      temp_ppt_file <- tempfile(fileext = ".pptx")
      print(ppt, target = temp_ppt_file)
      rv$ppt_file <- temp_ppt_file
      showNotification("PPT generation complete.", type = "message")
    })
  })
  
  # Download handler for the PPT
  output$download_ppt <- downloadHandler(
    filename = function() {
      "selected_trials.pptx"
    },
    content = function(file) {
      req(rv$ppt_file)
      file.copy(rv$ppt_file, file)
    }
  )
  
  # Preview of the dataframes for all categories
  output$slide_previews <- renderUI({
    req(rv$preview_data_list)
    preview_outputs <- lapply(names(rv$preview_data_list), function(name) {
      DT::dataTableOutput(paste0("preview_", make.names(name)))
    })
    do.call(tagList, preview_outputs)
  })
  
  # Render the previews for each category
  observe({
    req(rv$preview_data_list)
    for (name in names(rv$preview_data_list)) {
      local({
        name <- name  # Capture the current value of 'name'
        df_preview <- rv$preview_data_list[[name]]
        output_name <- paste0("preview_", make.names(name))
        output[[output_name]] <- DT::renderDataTable({
          DT::datatable(df_preview, options = list(scrollX = TRUE), caption = htmltools::tags$caption(
            style = 'caption-side: top; text-align: left; color: #007BFF; font-size: 150%; font-weight: bold; border-top: 1px solid #ddd; padding-top: 5px;', 
            paste("Preview for", name)
          ))
        })
      })
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)
