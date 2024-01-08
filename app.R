
library(shiny)
library(shinydashboard)
library(shinyalert)
library(shinyjs)
library(DT)
library(tidyverse)
library(dplyr)
library(rhandsontable)
library(tools)
library(officer)
library(docxtractr)
library(flextable)
library(mailR)
library(mefa)
library(shinybusy)
library(tidyr)
library(shinyWidgets)
options(repos = BiocManager::repositories())

set_libreoffice_path("/opt/libreoffice7.0/program/soffice")
options(shiny.maxRequestSize=30*1024^2)

labelMandatory <- function(label) {
    tagList(
        label,
        span("*", class = "mandatory_star")
    )}


download_btn <-  function (outputId, label = "Download", class = NULL, ...){
    tags$a(id = outputId, class = paste("btn btn-default shiny-download-link", 
                                        class), href = "", target = "_blank", download = NA, 
           icon("file-word"), label, ...)
}

download_pdf <-  function (outputId, label = "Download", class = NULL, ...){
    tags$a(id = outputId, class = paste("btn btn-default shiny-download-link", 
                                        class), href = "", target = "_blank", download = NA, 
           icon("file-pdf"), label, ...)
}

ui = 
    dashboardPage(skin = "black",
                  dashboardHeader(titleWidth='100%',
                                  tags$li(class = "dropdown",
                                          tags$style(".main-header {max-height: 80px}"),
                                          tags$style(".main-header .logo {height: 80px;padding: 0 0px;}"),
                                          actionLink("openModal", label = "", icon = icon("info-circle"))
                                  ),
                                  title = span(
                                      tags$img(src="affinia_crop.png", height = 80,width = 230),
                                      tags$img(src="affinia_color.png", height = 80,width = '1280vh', align = "right")
                                  )
                  ),
                  dashboardSidebar(width = 160,
                                   sidebarMenu(id="sbmenu",
                                               tags$head(tags$style(HTML(".left-side, .main-sidebar {padding-top: 100px;}"))
                                               ),
                                               menuItem("New Form", tabName = "Form", icon = icon("file-signature")),
                                               menuItem("Submitted Forms", tabName = "Submitted", icon = icon("clipboard-list")))
                  ),
                  dashboardBody(
                      
                      tabItems(
                          tabItem(tabName = "Form",
                                  fluidRow(id = "form",
                                           useShinyalert(),
                                           useShinyjs(),
                                           tags$head(
                                               tags$style(HTML("
                                               .navbar-nav>li>a {padding-top: 0px;padding-bottom: 5px;line-height: 0px;font-size: 36px;background-color: none;}
                                               .skin-black .main-header .navbar .nav>li>a {color: #5600ff;}
                                               .fa, .fab, .fad, .fal, .far, .fas {line-height: 0;}
                                               .shiny-split-layout > div {overflow: visible;}
                                               .handsontable {word-wrap: break-word;}
                                               .handsontable th {color: black !important; background: #00ce1f !important;font-weight: 600;font-size: 15px; }
                                               .handsontableEditor.autocompleteEditor, .handsontableEditor.autocompleteEditor .ht_master .wtHolder {min-height: 138px;}
                                               .handsontable.listbox tr td.current, .handsontable.listbox tr:hover td {background: #45d9f7;}
                                               .handsontable td.htInvalid { background-color: white !important; }
                                               .box {border: 3px solid #08de32;width: 150%; border-radius: 20px;padding: 30px;}
                                               .row {margin-right: 35px; margin-left: 50px;}
                                               .box-header h3.box-title {font-weight: bold;font-size: 24px;padding: 1px;}
                                               .col-sm-6 {width: 65%;}
                                               .table.dataTable tbody tr {word-wrap: break-word;}
                                               .input-group .form-control {color: #a22020;background-color: #fbd8d8;font-size: 14px; display: table-row;}
                                                               ")
                                               )
                                           ),
                                           box(
                                               
                                               title = HTML(paste(p(HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;')
                                                                    ,HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;'),HTML('&emsp;')
                                                                    ,HTML('&emsp;'),HTML('&emsp;'),strong("Animal Studies Form")))),
                                               radioButtons("type","STUDY TYPE",c("MR","MD","PR","PD"), inline = T),
                                               splitLayout(cellWidths = c("30%","4%","30%","4%","30%"),textInput("study_no",("Study ID"), width = '600px'), HTML(paste(HTML('&nbsp;'))),
                                                           dateInput("sub_date","Requested Date", value = Sys.time(),format = "yyyy-mm-dd"),HTML(paste(HTML('&nbsp;'))),
                                                           textInput("submitted_by",("Submitted By"), width = '600px')),
                                               selectInput("GLP", labelMandatory("GLP Required"),c("YES","NO"), selected = "NO", width = '600px'),
                                               h5("Funding Source: *Contact CBO if needed*",style = "font-weight: 800;font-size: 15px;"),
                                               splitLayout(cellWidths = c("40%","5%","40%"), textInput("dept","Department"),HTML(paste(HTML('&nbsp;'))),
                                                           textInput("proj","Project")),
                                               textInput("study_dir",labelMandatory("Study Director"), width = '600px'),
                                               textAreaInput("assoc_email", "Associate's email", width = '900px', resize = 'vertical'),
                                               textAreaInput("goal",labelMandatory("Study Goal and References"),width = '900px', height = '100px', resize = 'vertical'),
                                               radioButtons("design","Test Article Design",c("Clonal Study","Library Study","Others"), inline = T),
                                               textAreaInput("library","Test Article/Other Description",width = '900px', height = '100px', resize = 'vertical'),
                                               ###############################################################################################################################################################################
                                               
                                               h5("Experimental Design:",style = "font-weight: 800;font-size: 16px;"),
                                               textInput("species",("Animal Species"), width = '600px'),
                                               textInput("stock",("Stock #"), width = '600px'),
                                               numericInput("groups","Number of Groups",1,min = 1,width = '600px'),
                                               h5("Time Points for Each group:", style = "font-weight: 800;font-size: 16px;"),
                                               br(),
                                               rHandsontableOutput("timepoints_table"),
                                               br(),
                                               actionButton("save_timepoints", "SAVE TABLE",buttonLabel = "Browse",
                                                           style = "color: black; background-color:#45d9f7; border-color: #060606; font-size: 10px; font-weight: 1000;padding: 5px 10px;float:right;"),
                                               br(),
                                               
                                               ###############################################################################################################################################################################
                                               
                                               h5("Test Article (TA)",style = "font-weight: 800;font-size: 16px;"),
                                               splitLayout(cellWidths = c("35%","38%"),selectInput("manufacturer", labelMandatory("Manufacturer"),c("Affinia","Other"), selected = NULL),
                                                           textInput("manu_other","Other")),
                                               radioButtons("avail_date","Estimated Test Article Availability Date",c("Choose a date","TBD"), inline = T),
                                               dateInput("req_date","Request Date", value = Sys.time(), width = '600px',format = "yyyy-mm-dd"),
                                               p(id = "date_info", "Note: Study cannot be scheduled untill a date is provided."),
                                               textInput("conc",("Initial Concentration(if known)"), width = '600px'),
                                               fileInput("add_info","Additional Information",accept = c(".docx",".pptx",".doc", ".pdf"), width = '600px',placeholder = "Upload word documents or powerpoints only"),
                                               h5("Study Design Table:",style = "font-weight: 800;font-size: 16px;"),
                                               rHandsontableOutput("design_table"),
                                               br(),
                                               textAreaInput("endpoint_sample",labelMandatory("Endpoint Sample Collection"),width = '900px', height = '100px', resize = 'vertical'),
                                               p(id = "endpoint_info", "Note: Please list all the tissues that needs to be collected."),
                                               radioButtons("analysis_info","Analysis Developed in-house",c("Yes","No","I don't know"), inline = T),
                                               h5("Analysis Table:",style = "font-weight: 800;font-size: 16px;"),
                                               rHandsontableOutput("endpoints_table"),
                                               br(),br(),
                                               textAreaInput("notes","Extra Information",width = '900px', height = '100px', resize = 'vertical'),
                                               actionButton("submit", "SUBMIT FORM",buttonLabel = "Browse",
                                                            style = "color: black; background-color:#45d9f7; border-color: #060606; font-weight: 1000;float:right;")
                                               
                                               ###############################################################################################################################################################################
                                           )
                                  )
                          ),
                          ####################################################################################################################################################################################################
                          
                          # Second tab content
                          tabItem(tabName = "Submitted",
                                  div(DTOutput("requested_table"), style = "font-size:90%; font-weight: 500;")      
                          )
                          
                          ####################################################################################################################################################################################################
                  )
    )
    )


# Define server logic required to draw a histogram
server <- function(input, output, session) {
    research_team <- c("ssiripurapu","bmastis","rcalcedo")
##################################################################################################################################################################
    
    report_file <- reactive({
        file <- read.csv("/mnt/efs/prj/animal_studies_requests/requests_form.csv",stringsAsFactors = FALSE)
        file
    })
 
###################################################################################################################################################################
    #Conditional boxes manipulation
    
    observe({
        file <- report_file()
        file <- read.csv("/mnt/efs/prj/animal_studies_requests/requests_form.csv")
        if (input$type == "PR" | input$type == "PD"){
          file_firstrow <- file %>% filter(Study_Type == "PR" | Study_Type == "PD") %>%
          filter(row_number()==1)
        } else{
          file_firstrow <- file %>% filter(Study_Type == "MR" | Study_Type == "MD") %>%
          filter(row_number()==1)
        }
        study_num <- unlist(strsplit(file_firstrow$Study_ID,'-'))
        study_num <- sprintf("%04d", (as.numeric(study_num[3]) + 1))
        study_id <- paste0("AFT-",input$type,"-", study_num )
        updateTextInput(session, "study_no", value = study_id)
        disable("study_no")
    })
    
    user_name <- reactive({
        name <- session$user
    })
    

    observeEvent(session$user,{
        if (session$user %in% research_team){
            updateTabItems(session, "sbmenu", selected = "Submitted")
        }
        else{
            updateTabItems(session, "sbmenu", selected = "Form")
        }
    })
    
    observe({
        disable("manu_other")
        disable("sub_date")
        updateTextInput(session, "submitted_by", value = user_name())
        disable("submitted_by")
        updateTextInput(session, "study_dir", value = user_name())
        updateTextAreaInput(session, "assoc_email", value = paste0(user_name(),"@affiniatx.com,"))
    })
    
    
    observe({
        if (input$manufacturer == "Other"){
            shinyjs::enable("manu_other")
        } else {
            disable("manu_other") 
        }
    })
    observe({
        if (input$avail_date == "TBD"){
            shinyjs::hide("req_date")
            shinyjs::show("date_info")
        }
        if (input$avail_date == "Choose a date"){
            shinyjs::show("req_date")
            shinyjs::hide("date_info")
        }
    })
    
    observe({
        if (input$design == "Clonal Study"){
            shinyjs::disable("library")
            updateTextAreaInput(session, "library", value = "Library information is not relavant to this study design")
        }
        else {
           shinyjs::enable("library")
            updateTextAreaInput(session, "library", value = "")
        }
    })
    
    
    observeEvent(input$add_info,{
        if(!is.null(input$add_info))
            file <- input$add_info
            ext <- tools::file_ext(file$datapath)
            file_type <- c("docx","pptx", "pdf")
            if (!(ext %in% file_type)){
                shinyalert::shinyalert("Upload only Word or Powerpoint documents!", type = "error", closeOnClickOutside = T, timer = 2000, showConfirmButton = F) 
                reset("add_info")
            }
    })
    
    observe({
        if(input$GLP!="" && input$study_dir!="" &&input$goal!="" &&input$endpoint_sample!="") {
            enable("submit")
            }
        else {
            disable("submit")
            }
        })
    
    observeEvent(input$date_range_button, {
        print(input$date_range)
        #reset("date_range")
    })
    
    
###################################################################################################################################################################
    
    #Timepoints Table server processing:
    color_renderer <- "function(instance, td) {Handsontable.renderers.TextRenderer.apply(this, arguments);td.style.background = '#c3c0c0'; td.style.color = 'black'}"
    
    output$timepoints_table <-  renderRHandsontable({
            group_no <- input$groups
            time_df <- data.frame("Group_No." = as.numeric(),"No._of_Timepoints" = as.numeric(),"No. of Animals" = as.numeric(),"Animal's sex" = as.character(), 
                                  "Vendor" = as.character(),"Route of Adm." = as.character(),
                                  "Study Duration (pref in days)" = as.character(), "Age required at the start of study" = as.character(), check.names = FALSE)
            time_df[nrow(time_df)+group_no,] <- ""
            time_df$Group_No. <- rownames(time_df)
            time_df$No._of_Timepoints <- as.numeric(1)
            colnames(time_df) <- gsub("_"," ", colnames(time_df))
            
        rhandsontable(time_df, width = 1100,stretchH = "all",rowHeaders = NULL) %>%
            hot_cols(colWidths = c(20,20,40,20,20,20,30,30),manualColumnResize = TRUE) %>%
            hot_rows(rowHeights = 35) %>%
            hot_col("No. of Timepoints", readOnly = TRUE,renderer = color_renderer) %>%
            hot_col(col = "Route of Adm.", type = "dropdown", source = c("IV","IM","SQ","IP","PO","LP/IT","ICM"),allowInvalid = T) %>%
            hot_col(col = "Animal's sex", type = "dropdown", source = c("Male","Female"),allowInvalid = T) %>%
            hot_col(col = "Vendor", type = "dropdown", source = c("Jackson [JAX]","Charles River [CRL]","Taconic"),allowInvalid = T)
    })

    
    observe({
        if (!is.null(input$timepoints_table)){
            shinyjs::enable("save_timepoints")
        }
    })
    
    vars <- reactiveValues(counter = 0)
    
    observeEvent(input$save_timepoints,{
        if (!is.null(isolate(input$timepoints_table))) {
            x <- hot_to_r(isolate(input$timepoints_table))
            x[x==""]<-NA
            na_val <- any(is.na(x))
            if (na_val == TRUE){
                vars$counter <- 0
                shinyalert::shinyalert("Please fill all the Information", type = "error", closeOnClickOutside = T, timer = 3000, showConfirmButton = F)  
            } else {
                vars$counter <- vars$counter+1
                colnames(x) <- gsub(" ","_",colnames(x))
                temp_filepath <- paste0(tempdir(), "/timepoints.csv")
                write.csv(x, temp_filepath , row.names=FALSE)
                shinyalert::shinyalert("Saved Table", type = "success", closeOnClickOutside = T, timer = 3000, showConfirmButton = F)
                shinyjs::disable("save_timepoints")
            }
        }
        })
    
    
    ####################################################################################################################################################################
    
    #Study Design Table
    
    color_renderer <- "function(instance, td) {Handsontable.renderers.TextRenderer.apply(this, arguments);td.style.background = '#c3c0c0'; td.style.color = 'black'}"
    
    output$design_table <- renderRHandsontable({
        group_no <- input$groups
        if(vars$counter > 0) {
            temp_filepath <- paste0(tempdir(), "/timepoints.csv")
            timepoint_table <- read.csv(temp_filepath)
            design_df <- timepoint_table[,c("Group_No.","No._of_Animals","Route_of_Adm.")]
            namevector <- c("Strain","Treatment", "Dosage","Dosage Units","Time Post Procedure (pref in days)")
            design_df[ ,namevector] <- ""
        } else {
            design_df = data.frame("Group_No." = as.numeric(), "No._of_Animals" = as.character(), "Route_of_Adm." = as.character(),"Strain" = as.character(),
                                   "Treatment" = as.character(),"Dosage" = as.character(),"Dosage Units" = as.character(),
                                   "Time Post Procedure (pref in days)" = as.character(),check.names = FALSE)
        }
        
        colnames(design_df) <- gsub("_"," ", colnames(design_df))
        
        rhandsontable(design_df, width = 1100,stretchH = "all",rowHeaders = NULL) %>%
            hot_cols(colWidths = c(20,20,20,25,30,20,20,25),manualColumnResize = TRUE) %>%
            hot_col("Group No.", readOnly = TRUE,renderer = color_renderer) %>%
            hot_col("No. of Animals", readOnly = TRUE,renderer = color_renderer) %>%
            hot_col("Route of Adm.", readOnly = TRUE,renderer = color_renderer) %>%
            hot_col("Dosage Units", type = "dropdown", source = c("vg/kg","vg/animal","vg/gm of brain wt."),allowInvalid = T) %>%
            hot_rows(rowHeights = 35)
    })
    
    ##################################################################################################################################################################
    
    #Endpoints Table:
    
    color_renderer <- "function(instance, td) {Handsontable.renderers.TextRenderer.apply(this, arguments);td.style.background = '#c3c0c0'; td.style.color = 'black'}"
    
    
    output$endpoints_table <- renderRHandsontable({
       if(vars$counter > 0) {
         temp_filepath <- paste0(tempdir(), "/timepoints.csv")
         timepoint_tab <- read.csv(temp_filepath)
         timepoint_tab <-  timepoint_tab[,c("Group_No.","No._of_Timepoints")]
         timepoints <- timepoint_tab[!is.na(as.numeric(as.character(timepoint_tab$No._of_Timepoints))),]
         timepoints$No._of_Timepoints <- as.numeric(timepoints$No._of_Timepoints)
         
         timepoints <- timepoints %>% 
             group_by(Group_No.) %>%
             complete(No._of_Timepoints = full_seq(1:No._of_Timepoints,1))
         
         timepoints$No._of_Timepoints <- as.character(timepoints$No._of_Timepoints)
         colnames(timepoints)[[2]] <- "Timepoint No."
         new_cols <- c("Analysis Timepoint", "Tissue", "Location of Analysis", "Analysis Type")
         timepoints[new_cols] <- ""
         endpoint_df <- timepoints
         
       } else {
           endpoint_df <- data.frame("Group_No." = as.numeric(),"Timepoint No." = as.numeric(),"Analysis Timepoint" = as.character(),"Tissue" = as.character(), 
                                     "Location of Analysis" = as.character(), "Analysis Type" = as.character(), check.names = FALSE)
           endpoint_df[nrow(endpoint_df)+1,] <- ""
       }
        
        #tissues_mentioned <- input$endpoint_sample
        colnames(endpoint_df) <- gsub("_"," ", colnames(endpoint_df))
        
        
        rhandsontable(endpoint_df, width = 1100,stretchH = "all",rowHeaders = NULL,highlightRow = FALSE, highlightCol = FALSE,selectCallback = TRUE) %>%
            hot_cols(colWidths = c(20, 20,30,60,40,60),manualColumnResize = TRUE) %>%
            hot_col("Timepoint No.", readOnly = FALSE,renderer = color_renderer, default = 1) %>%
            hot_col("Group No.", readOnly = FALSE,renderer = color_renderer) %>%
            hot_rows(rowHeights = 30 )
        
    })

    
    ##################################################################################################################################################################
    #Capturing the inputs for appending to excel
    Inputs <- reactive({
        
        assoc_email <- input$assoc_email
        assoc_email <- gsub(" ","",assoc_email)
        assoc_email <- str_replace(assoc_email,",$","")
        original_em <- assoc_email
        if (input$design == "Library Study"){
            assoc_email <- paste0("animalstudy@affiniatx.com,Computational_Science@affiniatx.com,afieldsend@affiniatx.com,",assoc_email, sep = "") 
        } else {
            assoc_email <- paste0("animalstudy@affiniatx.com,",assoc_email, sep = "")  
        }
        
        original_email <- paste0("ssiripurapu@affiniatx.com,bmastis@affiniatx.com,rcalcedo@affiniatx.com,",original_em, sep = "")
        original_path <- paste0("/mnt/efs/prj/animal_studies_requests/",input$type,"/original_documents/",input$study_no,"_original.docx", sep = "")
        uploaded_path <- paste0("/mnt/efs/prj/animal_studies_requests/",input$type,"/reviewed documents/",input$study_no,"_reviewed.pdf", sep = "")
        req_date <- (as.character(input$sub_date))
        
        if(is.null(input$add_info)){
           additional_info <- "No document was submitted"
           add_detail <-"No Additional Document is submitted"
        } else{
            additional_info <- paste0("/mnt/efs/prj/animal_studies_requests/",input$type,"/additional_documents/",input$study_no,"_additional.pdf",sep = "")
            add_detail <-"Additional Document is submitted. Please download from the app"
        }
        
        inputs_df <- data.frame(input$type, input$study_no, req_date, input$submitted_by,assoc_email,original_path,additional_info,"Under Review","","",uploaded_path )
        combo <- list(inputs_df = inputs_df, add_detail = add_detail, assoc_email = assoc_email,original_email = original_email)
        combo
    })
        
    
    ##############################################################################
    
    #Rendering Information into word document
    
    observeEvent(input$submit,{
        filepath <- paste0("/mnt/efs/prj/animal_studies_requests/",input$type,"/original_documents/",input$study_no,"_original.docx", sep = "")
        ######
        
        sub_title <- paste0("\t Study No - ",input$study_no, sep = "")
        Funding_Source <- paste0(input$dept,"-", input$proj, sep = "")
        if (input$avail_date == "Choose a date"){
            endpoint_date <- (as.character(input$req_date))
        } else {
            endpoint_date <- "TBD"
        }
        sub_date <- (as.character(input$sub_date))
        #######
        
        ##Affinia image to center
        img_in_par <- fpar(
            external_img(src = "/mnt/efs/prj/rsconnect_rshiny/Animal_Studies/www/affinia_crop.png", height = 0.8, width = 4.2),
            fp_p = fp_par(text.align = "center",padding.top = 85))
        ######
        
        title_prop <- fp_text(color = "black",font.size = 14,bold = T,font.family = "Cambria (Body)")
        par_style <- fp_par(text.align = "center", padding = 4)
        par_style1 <- fp_par(text.align = "center", padding = 4, border.bottom = fp_border(color = "#45d9f7", style = "solid", width = 2))
        ######
        sub_text <- fp_text(color = "black",font.size = 12,bold = T,font.family = "Times New Roman")
        sub_align <- fp_par(text.align = "left", padding = 4,line_spacing = 2)
        ######
        
        headline_text <- fp_text(color = "black",font.size = 12,bold = T,font.family = "Times New Roman")
        headline_align <-  fp_par(padding = 5)
        
        regular_text <- fp_text(color = "black",font.size = 12,bold = F,font.family = "Times New Roman")
        #### timepoints table
        timepoints_DF <- hot_to_r(input$timepoints_table)
        ft_time <- flextable(timepoints_DF) %>% 
            width(~Vendor, 1) %>%
            #width(~Age_required_at_the_start_of_study, 1) %>%
            border( border = fp_border(color="black"), part = "all" ) %>%
            fontsize(size = 10, part = "body") %>%
            fontsize(size = 10, part = "header") %>%
            bold(part = "header") %>%
            align( align = "center", part = "all" ) %>%
            font(fontname = "Times New Roman", part = "all") %>%
            bg(bg = "#86d281", part = "header")
        #######Study Deisgn
        designDF <- hot_to_r(input$design_table)
        ft_design <- flextable(designDF) %>% 
            border( border = fp_border(color="black"), part = "all" ) %>%
            fontsize(size = 10, part = "body") %>%
            fontsize(size = 10, part = "header") %>%
            bold(part = "header") %>%
            align( align = "center", part = "all" ) %>%
            font(fontname = "Times New Roman", part = "all") %>%
            bg(bg = "#86d281", part = "header")
        ###Analysis Table
        analysisDF <- hot_to_r(input$endpoints_table)
        ft_analysis <- flextable(analysisDF) %>%
            width(width = 1) %>%
            border( border = fp_border(color="black"), part = "all" ) %>%
            fontsize(size = 10, part = "body") %>%
            fontsize(size = 10, part = "header") %>%
            bold(part = "header") %>%
            align( align = "center", part = "all" ) %>%
            font(fontname = "Times New Roman", part = "all") %>%
            bg(bg = "#86d281", part = "header")
        ##########
        
        #Making the document
        doc <- officer::read_docx() %>% 
            officer::body_add_fpar(img_in_par) %>%
            officer::body_add_fpar(fpar(ftext("\t In-Vivo Study Submission Request Form", prop = title_prop), fp_p = par_style )) %>%
            officer::body_add_fpar(fpar(ftext(sub_title, prop = title_prop), fp_p = par_style )) %>%
            officer::body_add_fpar(fpar(ftext("", prop = title_prop), fp_p = par_style1 )) %>%
            officer::body_add_fpar(fpar(ftext(""))) %>%
            
            officer::body_add_fpar(fpar(ftext("1. STUDY DETAILS:", prop = sub_text))) %>%
            officer::body_add_fpar(fpar(ftext("\t a. Submitted On:\t", prop = headline_text),ftext(sub_date, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t b. Submitted By:\t", prop = headline_text),ftext(input$submitted_by, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t c. GLP Required:\t", prop = headline_text),ftext(input$GLP, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t d. Funding Source:\t", prop = headline_text),ftext(Funding_Source, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t e. Study Director:\t", prop = headline_text),ftext(input$study_dir, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t f. Associate's email:\t", prop = headline_text),ftext(Inputs()$assoc_email, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t g. Study Goal and References:\t", prop = headline_text),ftext(input$goal, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t h. Test Article Design:\t", prop = headline_text),ftext(input$design, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t i. Test Article/Other Description:\t", prop = headline_text),ftext(input$library, prop = regular_text),fp_p = headline_align)) %>%
            
            officer::body_add_fpar(fpar(ftext("", prop = title_prop), fp_p = par_style1 )) %>%
            officer::body_add_fpar(fpar(ftext("2. EXPERIMENTAL DESIGN:", prop = sub_text))) %>%
            officer::body_add_fpar(fpar(ftext("\t a. Animal Species:\t", prop = headline_text),ftext(input$species, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t b. Stock #:\t", prop = headline_text),ftext(input$stock, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t c. Number of Groups:\t", prop = headline_text),ftext(input$groups, prop = regular_text),fp_p = headline_align)) %>%
            
            officer::body_add_fpar(fpar(ftext("Timepoints Table:", prop = sub_text))) %>%
            body_add_flextable(ft_time) %>%
            
            officer::body_add_fpar(fpar(ftext("", prop = title_prop), fp_p = par_style1 )) %>%
            officer::body_add_fpar(fpar(ftext("3. TEST ARTICLE:", prop = sub_text))) %>%
            officer::body_add_fpar(fpar(ftext("\t a. Manufacturer:\t", prop = headline_text),ftext(input$manufacturer, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t b. Estimated Test Article Availability Date:\t", prop = headline_text),ftext(endpoint_date, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t c. Initial Concentration:\t", prop = headline_text),ftext(input$conc, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t c. Additional Information:\t", prop = headline_text),ftext(Inputs()$add_detail, prop = regular_text),fp_p = headline_align)) %>%
            
            officer::body_add_fpar(fpar(ftext("Study Design Table:", prop = sub_text))) %>%
            body_add_flextable(ft_design) %>%
            
            officer::body_add_fpar(fpar(ftext("", prop = title_prop), fp_p = par_style1 )) %>%
            officer::body_add_fpar(fpar(ftext("4. TISSUE ANALYSIS:", prop = sub_text))) %>%
            officer::body_add_fpar(fpar(ftext("\t a. Endpoint Sample Collection:\t", prop = headline_text),ftext(input$endpoint_sample, prop = regular_text),fp_p = headline_align)) %>%
            officer::body_add_fpar(fpar(ftext("\t b. Analysis Developed in-house:\t", prop = headline_text),ftext(input$analysis_info, prop = regular_text),fp_p = headline_align)) %>%
            
            officer::body_add_fpar(fpar(ftext("Analysis Table:", prop = sub_text))) %>%
            body_add_flextable(ft_analysis) %>%
            
            officer::body_add_fpar(fpar(ftext("Extra Information:\t", prop = headline_text),ftext(input$notes, prop = regular_text),fp_p = headline_align))
        
        print(doc, target = filepath)
        #######################
        
        #writing the new information to excel
        original_df <- report_file()
        new_df <- Inputs()$inputs_df
        colnames(new_df) <- colnames(original_df)
        total_df <- rbind.data.frame(new_df,original_df)
        write.csv(total_df, file="/mnt/efs/prj/animal_studies_requests/requests_form.csv", row.names = F)
        
        ##Saving the additional documents
        if(!is.null(input$add_info)) {
            ext <- tools::file_ext(input$add_info$datapath)
            uploaded_addfile_Path <- input$add_info$datapath
            additional_path <- paste0("/mnt/efs/prj/animal_studies_requests/",input$type,"/additional_documents/",input$study_no,"_additional.pdf", sep = "")
            if (ext == "pdf") {
              file.copy(from = uploaded_addfile_Path, to = additional_path ) 
            } else{
              pdf <- convert_to_pdf(uploaded_addfile_Path, pdf_file = additional_path)
            }
        }
        
        #####
        row_data <- Inputs()$inputs_df
        body <- paste0("A new animal study request with Study Id: ", row_data[,2], " by ",row_data[,4], " was submitted. A copy of the submitted document is attached to this email.")
        email_to <- Inputs()$original_email
        email_to <- unlist(strsplit(email_to, ","))
        email_to <- dput(as.character(email_to))
        send.mail(from = "noreply@affiniatx.com",
                  to = email_to,
                  subject = "New Notification on Animal Study Form",
                  body = body,
                  authenticate = TRUE,
                  attach.files = row_data[,6],
                  smtp = list(host.name = "smtp.office365.com",
                              port = 587,
                              user.name = "noreply@affiniatx.com",
                              passwd = "Tuh71036",
                              tls = TRUE))
        
        
        shinyalert::shinyalert("Succesfully Submitted", type = "success", closeOnClickOutside = T, timer = 3000, showConfirmButton = F)
        Sys.sleep(1.5)
        session$reload()
        
        
    })
    
    ###################################################################################################################################################################
    ###Submitted datatable
    
    buttonInput <- function(data_file,FUN, len, id, ...) {
        inputs <- character(len)
        for (i in seq_len(len)) {
            inputs[i] <- as.character(FUN(paste0(id, i), ...))
            
        }
        inputs
    }
    
    ###################################################################################################################################################################
    
    submitted_table <- reactive({
        data_file <- report_file()
        data_file2 <- report_file()
        data_file$Original_Form <- buttonInput(
            data_file <- data_file,
            FUN = download_btn,
            len = nrow(data_file),
            id = 'button_',
            label = "",
            style = "border: none;background: none;font-size:170%;color: #071aff;",
            onclick = ('Shiny.setInputValue("originalData", this.id)')
        )

        data_file$Reviewed_Form <- buttonInput(
            data_file <- data_file,
            FUN = download_pdf,
            len = nrow(data_file),
            id = 'buttn_',
            label = "",
            style = "border: none;background: none;font-size:170%;color: #ff0707;",
            onclick = ('Shiny.setInputValue("ReviewedData", this.id)')
        )
        
        data_file$Upload_Reviewed <- buttonInput(
            data_file <- data_file,
            FUN = actionButton,
            len = nrow(data_file),
            id = 'butn_',
            label = "",
            icon = icon("file-upload"),
            style = "border: none;background: none;font-size:170%;color: #10c3ca;",
            onclick = ('Shiny.setInputValue("uploadData", this.id)')
        )
        
        data_file$Additional_Info <- buttonInput(
            data_file <- data_file,
            FUN = download_pdf,
            len = nrow(data_file),
            id = 'btn_',
            label = "",
            style = "border: none;background: none;font-size:170%;color: #ff0707;",
            onclick = ('Shiny.setInputValue("AddData", this.id)')
        )
        
        
        for (row in 1:nrow(data_file)){
            if (data_file[row,8] != "Reviewed"){
                data_file[row,9] <- "No Document"
            }
        }
        
        for (row in 1:nrow(data_file2)){
            if ((data_file2[row,7]) == "No document was submitted"){
                data_file[row,7] <- "No document was submitted"
            } 
        }
        
        data_file$Additional_Info <- gsub("No document was submitted",'<strong style="color:black">No document was submitted</strong>',data_file$Additional_Info)
        data_file$Upload_Reviewed <- gsub("Doc Uploaded",'<strong style="color:green">Doc Uploaded</strong>',data_file$Upload_Reviewed)
        data_file$Reviewed_Form <- gsub("No Document",'<strong style="color:black">No Document</strong>',data_file$Reviewed_Form)
        data_file$Review_status <- gsub("Reviewed",'<strong style="color:green">Reviewed</strong>',data_file$Review_status)
        data_file$Review_status <- gsub("Under Review",'<strong style="color:red">Under Review</strong>',data_file$Review_status)
        data_file
    })
   
    ############################
    #File Upload in server table
    
    upload_dataModal <- function(failed = FALSE) {
        modalDialog(
            tags$h2('Upload reviewed document'),
            fileInput("reviewed_doc","Upload Document",accept = c(".doc",".docx"), width = '600px',placeholder = "Please upload microsoft word documents only"),
            footer=tagList(
                actionButton('upload_submit', 'Submit'),
                modalButton('Cancel')
            )
        )
    }
    
    observeEvent(input$uploadData,{
        showModal(upload_dataModal())
    })
    
    observeEvent(input$upload_submit,{
        
        File <- input$reviewed_doc
        file_tag <- file_ext(File$name)
        file_type <- c("doc","docx")
        
        if (!is.null(File) && file_tag %in% file_type ) {
            removeModal()
            selected_row <- as.numeric(gsub("butn_","",input$uploadData))
            row_data <- report_file()[selected_row,]
            upload_to <- row_data[,"Upload_Reviewed"]
            uploadedfileDataPath<- input$reviewed_doc$datapath
            
            status <- row_data[,"Review_status"]
            
            if (status != "Reviewed"){
                doc_file <- convert_to_pdf(uploadedfileDataPath,upload_to)
            } else {
                file.remove(upload_to)
                doc_file <- convert_to_pdf(uploadedfileDataPath,upload_to)
            }
            
            #### Email notification
            email_body <- paste0("Animal study form with Study ID :",row_data[,2]," which was submitted by ",row_data[,4], " on ", row_data[,3],
                                 " was reviewed and the document is uploaded. For quick review the reviewed document is attached to this email.")
            email_sent <- row_data[,5]
            email_sent <- unlist(strsplit(email_sent, ","))
            email_sent <- dput(as.character(email_sent))
            send.mail(from = "noreply@affiniatx.com",
                      to = email_sent,
                      subject = "New Notification on Animal Study Form",
                      body = email_body,
                      authenticate = TRUE,
                      attach.files = upload_to,
                      smtp = list(host.name = "smtp.office365.com",
                                  port = 587,
                                  user.name = "noreply@affiniatx.com",
                                  passwd = "Tuh71036",
                                  tls = TRUE))
            
            input_file <- report_file()
            input_file[selected_row,8] <- "Reviewed"
            input_file[selected_row,9] <- report_file()[selected_row,11]
            input_file[selected_row,10] <- as.character(format(Sys.Date(), "%m/%d/%Y"))
            write.csv(input_file, "/mnt/efs/prj/animal_studies_requests/requests_form.csv", row.names = F)
            shinyalert::shinyalert("Upload Successful", type = "success", closeOnClickOutside = T, timer = 2000, showConfirmButton = F)
            Sys.sleep(1.5)
            session$reload()
            ####
            
        } else {
            shinyalert::shinyalert("Upload only Word documents!", type = "error", closeOnClickOutside = T, timer = 2000, showConfirmButton = F)
            showModal(upload_dataModal(failed = TRUE))
        }
        
    })
    
    #######################
    #Download files
    
    report_data <- read.csv("/mnt/efs/prj/animal_studies_requests/requests_form.csv")
    lapply(1:nrow(report_data), function(i){
        documentpath <- report_data[i,6]
        study_id <- report_data[i,2]
        output[[paste0("button_",i)]] <- downloadHandler(
            filename = function(){
                return(paste0(study_id,"_original.docx"))
            },
            content = function(file){
                file.copy(documentpath, file)
            },
            contentType = "text"
        )
    })
    
    report_data_pdf <- read.csv("/mnt/efs/prj/animal_studies_requests/requests_form.csv")
    lapply(1:nrow(report_data_pdf), function(i){
        pdf_path <- report_data[i,9]
        studyid <- report_data[i,2]
        output[[paste0("buttn_",i)]] <- downloadHandler(
            filename = function(){
                return(paste0(studyid,"_reviewed.pdf"))
            },
            content = function(file){
                file.copy(pdf_path, file)
            },
            contentType = "pdf"
        )
    })
    
    report_data_add <- read.csv("/mnt/efs/prj/animal_studies_requests/requests_form.csv")
    lapply(1:nrow(report_data_add), function(i){
        add_path <- report_data[i,7]
        study_add <- report_data[i,2]
        output[[paste0("btn_",i)]] <- downloadHandler(
            filename = function(){
                return(paste0(study_add,"_additional.pdf"))
            },
            content = function(file){
                file.copy(add_path, file)
            },
            contentType = "pdf"
        )
    })
    
    ###################################################################################################################################################################
    
    report_file_user <- reactive({
        if (session$user %in% research_team) {
            data <- submitted_table()
        } else {
            data <- submitted_table()[,-11]
        }
        data
    })
    
    
    output$requested_table <- DT::renderDataTable({
        data <- report_file_user()
        data$Associates_email <- gsub(",",", ",data$Associates_email)
        colnames(data) <- gsub("_"," ", colnames(data))
        data
        
    }, escape =F,server=F,class = 'cell-border stripe',options = list(ordering=F,paging = F,selection=list(mode="single"),
                                                                      columnDefs = list(list(
                                                                                        className = 'dt-center', targets = "_all",
                                                                                        width = '200px', targets = "2")),
                                                                      preDrawCallback = JS('function() { Shiny.unbindAll(this.api().table().node()); }'),
                                                                      drawCallback = JS('function() { Shiny.bindAll(this.api().table().node()); } ')))
    
    ###############################################################################################################################################################
    #Help Modal Activity
    
    observeEvent(input$openModal, {
        showModal(
            modalDialog(
                actionButton('help_page', 'OPEN ANIMAL STUDIES DESCRIPTION DOCUMENT',
                             onclick ="window.open('https://tdtx.sharepoint.com/:w:/s/TDTxScience/EVV5ZyDAkWBGuUwnZhGkP-IB60DY71VXsugKcsq7-TiYjg?e=gh58hV', '_blank')",
                             style = "color: black; background-color:#f36e6e; border-color: #060606; font-size: 15px; font-weight: 1000;"),
                footer = tagList(
                    h5("For Queries about the Application contact BRYAN MASTIS (bmastis@affiniatx.com)",style = "font-weight: 800;font-size: 14px;text-align: left;margin-top: 1px;"),
                    modalButton("Cancel"),
                ),
                easyClose = TRUE
                )
        )
    })
    
    observeEvent(input$help_page,{
        Sys.sleep(2)
        removeModal()
    })
    
    ###########################################################################################################################################################
    
    session$onSessionEnded(function() {
        temp_filepath <- paste0(tempdir(), "/timepoints.csv")
        if (file.exists(temp_filepath)) {
            file.remove(temp_filepath)
        }
    })
}

# Run the application 
shinyApp(ui = ui, server = server)
