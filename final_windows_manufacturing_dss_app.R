library(shiny)
library(shinydashboard)
library(ggplot2)
library(rlang)
library(readxl)
library(tidyverse)
library(caret)
library(DT)
library(broom)
library(pROC)
library(lpSolve)

# ---- Data Loading and Preparation ----

df <- read_excel("Window_Manufacturing.xlsx",
                 col_names = c("breakage_rate", "window_size", "glass_thickness",
                               "ambient_temp", "cut_speed", "edge_deletion_rate",
                               "spacer_distance", "window_color", "window_type",
                               "glass_supplier", "silicon_viscosity", "supplier_location"),
                 skip = 1
) %>%
  mutate(pass_fail = factor(ifelse(breakage_rate < 7, "Pass", "Fail")),
         across(where(is.character), as.factor),
         across(where(is.numeric), as.numeric))

data <- df

data_dictionary <- data.frame(
  Variable = c("breakage_rate", "pass_fail", "window_size", "glass_thickness",
               "ambient_temp", "cut_speed", "edge_deletion_rate",
               "spacer_distance", "window_color", "window_type",
               "glass_supplier", "silicon_viscosity", "supplier_location"),
  "Data Type" = c("numeric", "categorical", "numeric", "numeric", 
                  "numeric", "numeric", "numeric", "numeric", "numeric",
                  "categorical", "categorical", "numeric", "categorical"),
  "Business Use" = c("predictive target", "predictive target", "customer specification",
                     "customer specification", "manufacturer specification", "manufacturer specification",
                     "manufacturer specification", "manufacturer specification", "customer specification",
                     "customer specification", "manufacturer specification", "manufacturer specification",
                     "manufacturer specification"),
  "Model Use" = c("target", "target", rep("input", times = 11)),
  "Defintion" = c(
    "The average number of windows broken per batch",
    "A binary feature indicating if the batch met breakage specifications or not",
    "Diagonal of the recetangle in inches",
    "Thickness of the glass in inches",
    "The ambient temperate in Celcius when window is cut (0 degrees F to 75.2 F)",
    "A CNC machine cut in meters per minute",
    "A CNC machine edging in meters per minute",
    "Distance between brace spacers in inches",
    "Visible Light Transmittance (VLT) percentage is percentage of visible light transmitted through the window glass. Low values indicate darker glass, while larger numbers indicate more clear glass",
    "Window frame made of: Aluminum, Vinyl, or Wood",
    "Supplier A, Supplier B, Supplier C, Supplier D",
    "The viscosity of the silicone sealants or adhesives used in the assembly or installation of the window.  A crucial property that affects their application and performance. Larger number implies greater thickness",
    "Iowa, Michigan, Minnesota, Wisconsin"
  ),
  check.names = FALSE
) %>% 
  arrange(Variable)

# ---- Descriptive Plots ----

breakage_rate_mean <- data.frame(mean = round(mean(df$breakage_rate), 2))

plot_histogram <- ggplot(df) +
  geom_histogram(aes(x = breakage_rate), color = "white", fill = "steelblue4") +
  geom_vline(data = breakage_rate_mean, aes(xintercept = mean),
             linetype = "dashed", color = "black", size = 1, show.legend = FALSE) +
  geom_label(data = breakage_rate_mean,
             aes(x = mean + 0.5, y = 200, label = paste("Mean:", mean)), hjust = 0) +
  scale_y_continuous(expand = c(0,0)) +
  scale_x_continuous(breaks = seq(from = 0, to = 40, by = 5)) +
  labs(y = "Frequency", x = "Breakage Rate") +
  theme_classic() +
  theme(axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.title.x = element_text(size = 18, margin = margin(t = 10)),
        axis.text = element_text(size = 16)
  )

supplier_mean <- subset(df, !is.na(glass_supplier)) %>%
  group_by(glass_supplier) %>%
  summarise(mean = round(mean(breakage_rate),2)) %>%
  mutate(annotate = mean + 2)

plot_supplier <- subset(df, !is.na(glass_supplier)) %>%
  ggplot(aes(x = glass_supplier, y = breakage_rate)) +
  geom_jitter(aes(color = glass_supplier), show.legend = FALSE, alpha = 0.5, size = 2) +
  stat_summary(fun = mean, geom = "crossbar", color = "black", size = 0.5) +
  geom_label(data = supplier_mean,
             aes(x = glass_supplier, y = annotate, fill = glass_supplier, 
                 label = paste("Mean:",mean)), show.legend = FALSE) +
  labs(x = NULL, y = "Breakage Rate") +
  theme_classic() +
  theme(axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.text.x = element_text(size = 18, color = "black", margin = margin(t = 5, b = 5)),
        axis.text = element_text(size = 16)
  )

plot_cut_speed <- ggplot(df, aes(x = cut_speed, y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 2) +
  geom_smooth() +
  labs(x = "Cut Speed (meters per minute)", y = "Breakage Rate") +
  theme_classic() +
  theme(axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.title.x = element_text(size = 18, margin = margin(t = 10)),
        axis.text = element_text(size = 16)
  )

plot_glass_thickness <- ggplot(df, aes(x = glass_thickness, y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 1.75) +
  geom_smooth() +
  labs(x = "Glass Thickness (inches)", y = "Breakage Rate") +
  theme_classic() +
  theme(axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.title.x = element_text(size = 18, margin = margin(t = 10)),
        axis.text = element_text(size = 16)
  )

plot_window_size <- ggplot(df, aes(x = window_size, y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 1.75) +
  geom_smooth() +
  labs(x = "Diagonal Window Size (inches)", y = "Breakage Rate") +
  theme_classic() +
  theme(axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.title.x = element_text(size = 18, margin = margin(t = 10)),
        axis.text = element_text(size = 16)
  )

# ---- Predictive Variables ----

controllable_vars <- c("cut_speed", "edge_deletion_rate", "spacer_distance", "silicon_viscosity")
uncontrollable_vars <- c(
  "ambient_temp", "window_size", "glass_thickness", "window_color",
  "window_type", "glass_supplier", "supplier_location"
)

# ---- UI ----

ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(title = "WindowCo Helper",
                  dropdownMenu(type = "messages", badgeStatus = "success",
                               messageItem("Michael Frank", "frank118@purdue.edu"),
                               messageItem("Kasturirangan Penugonde", "kpenugon@purdue.edu"),
                               messageItem("Jonathan Roberts", "rober802@purdue.edu"),
                               messageItem("Daniel Aveline", "daveline@purdue.edu")
                  )
  ),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Home", tabName = "home", icon = icon("house-chimney")),
      menuItem("Descriptive Analytics", tabName = "descriptive", icon = icon("table-list")),
      menuItem("Predictive Analytics", tabName = "predictive", icon = icon("magnifying-glass")),
      menuItem("Prescriptive Analytics", tabName = "prescriptive", icon = icon("pills")),
      menuItem("Source Code", icon = icon("code"),
               href = "https://github.com/krangan2587/window-manufacturing-dss-app")
    )
  ),
  dashboardBody(
    tabItems(
      tabItem(tabName = "home",
              fluidRow(
                box(width = 6,
                    h2(strong("Welcome to Window Wonder!")),
                    h4("This application helps you, our", strong("WindowCo Product Team member"),", to understand how different variables – manufacturing parameters, customer specifications, and the intersection of both – affect our ability to minimize breakage and keep our customers satisfied."),
                    h4("This app contains three tabs:"),
                    tags$ul(
                      tags$li(h4(strong("Descriptive analytics:"), "This allows you to explore numerical relationships between the full range of customer specifications and manufacturing specifications.")),
                      tags$li(h4(strong("Predictive analytics:"), "This allows you to analyze predictable relationships between specifications and their impact on key quality outcome indicators: Breakage Rate (average windows broken per batch) or Pass/Fail (binary for the total batch).")),
                      tags$li(h4(strong("Prescriptive analytics:"), "This extends predictive analytics to identify how WindowCo can make future business decisions that optimize quality of our product, improve customer satisfaction, and enhance profitability as a result."))
                    )
                ),
              ),
              fluidRow(
                box(width = 6,
                    h2(strong("Analytic problem framing")),
                    h4("The fundamental question is: what window parameters are most associated with breakage and batch pass/failure, and how can we change our business decisions to improve those metrics?"),
                    h4("Example insights could include:"),
                    tags$ul(
                      tags$li(h4("Large pane sizes with Iowa-supplied glass have a 50% higher breakage rate than smaller pane sizes – we should consider optimizing our glass supply mix across other locations.")),
                      tags$li(h4("Dark glass (i.e., low visible light transfer) cut at higher ambient temperatures has a significantly higher batch failure rate – we should examine how we can adjust factory planning to cut these batches at lower temperatures, or if this is a product we want to consider changing."))
                    ),
                    h4("Because our inputs contain customer-driven parameters and manufacturer-driven parameters, the application offers the user maximum flexibility to explore and describe variable interactions, isolate variable combinations that predict window quality outcomes, and prescribe optimal values."),
                    h4("For prediction specifically, we have offered the option to use linear/logistic regression, elastic net, and random forest predictive methods – because there are a range of variables that may or may not have interactive relationships. We find random forest offers the most useful insights because it is more discovery-oriented (requiring the user to make fewer assumptions).")
                )
              ),
              fluidRow(
                box(width = 15,
                    h2(strong("Data Dictionary")),
                    DTOutput('data_dictionary'))
              )
      ),
      tabItem(tabName = "descriptive",
              h1(strong("Descriptive Analysis")),
              h4("Please explore our data set with a few preset graphics or with our custom plot generator."),
              fluidRow(
                box(title = "Distribution of Breakage Rate", solidHeader = TRUE, status = "primary", plotOutput("histogram")),
                box(title = "Plot your own graph here!", solidHeader = TRUE, status = "info", plotOutput("custom_plot"))
              ),
              fluidRow(
                box(title = "Breakage Rate vs. Glass Supplier", solidHeader = TRUE, status = "primary", plotOutput("supplier")),
                box(title = "Custom Plot Inputs", solidHeader = TRUE, status = "info",
                    varSelectInput("y_variable", "Select your y variable:", data),
                    varSelectInput("x_variable", "Select your x variable:", data, selected = "ambient_temp"),
                    actionButton("updateGraph", "Update", icon = icon("arrows-rotate"))
                )
              ),
              fluidRow(box(title = "Breakage Rate vs. Cut Speed", solidHeader = TRUE, status = "primary", plotOutput("cut_speed"))),
              fluidRow(box(title = "Breakage Rate vs. Glass Thickness", solidHeader = TRUE, status = "primary", plotOutput("glass_thickness"))),
              fluidRow(box(title = "Breakage Rate vs. Window Size", solidHeader = TRUE, status = "primary", plotOutput("window_size")))
      ),
      tabItem(tabName = "predictive",
              h1(strong("Predictive Analytics")),
              sidebarLayout(
                sidebarPanel(
                  h3(strong("Model Settings")),
                  selectInput("response", "Choose Response Variable:", choices = c("breakage_rate", "pass_fail"), selected = "breakage_rate"),
                  uiOutput("model_type_ui"),
                  checkboxGroupInput("predictors", "Select Predictors:",
                                     choices = setdiff(colnames(data), c("breakage_rate", "pass_fail")),
                                     selected = setdiff(colnames(data), c("breakage_rate", "pass_fail"))
                  ),
                  checkboxGroupInput("preprocess", "Preprocessing Steps:",
                                     choices = c("Center" = "center", "Scale" = "scale", "Impute Missing (median)" = "medianImpute", "Remove Near Zero Var" = "nzv"),
                                     selected = c("center", "scale")
                  ),
                  checkboxInput("use_cv", "10-Fold Cross-Validation", value = TRUE),
                  selectInput("caret_method", "Model Type (caret):",
                              choices = c("Linear/Logistic Regression" = "glm", "Elastic Net (glmnet)" = "glmnet", "Random Forest" = "rf"),
                              selected = "glm"
                  ),
                  conditionalPanel(
                    condition = "input.caret_method == 'glmnet'",
                    sliderInput("alpha", "Elastic Net Mixing (alpha):", min = 0, max = 1, value = 0.5)
                  ),
                  actionButton("fit_model", "Fit Model", icon = icon("cogs"))
                ),
                mainPanel(
                  h4(strong("Model Coefficient Table / Variable Importance")),
                  DTOutput("coef_table"),
                  h4(strong("Model Fit Evaluation")),
                  plotOutput("model_fit_plot"),
                  uiOutput("residual_plot_box"),
                  verbatimTextOutput("fit_summary")
                )
              )
      ),
      tabItem(tabName = "prescriptive",
              h1(strong("Prescriptive Analytics (LPsolve)")),
              sidebarLayout(
                sidebarPanel(
                  h4(strong("Set Uncontrollable Variables:")),
                  uiOutput("presc_sliders"),
                  h4(strong("Set/Override Bounds for Controllables:")),
                  uiOutput("controllable_bounds_ui"),
                  h4(strong("Optimization Objective:")),
                  sliderInput("min_br_constraint", "Minimum Predicted Breakage Rate (constraint):",
                              min = 0,
                              max = max(data$breakage_rate, na.rm = TRUE) + 10,
                              value = 0.01,
                              step = 0.01),
                  tags$p("Minimize Breakage Rate (constrained above) using linear programming."),
                  actionButton("presc_optimize", "Optimize Actions", icon = icon("bolt"))
                ),
                mainPanel(
                  h4(strong("Model Coefficient Table (for optimization)")),
                  DTOutput("presc_coef_table"),
                  h4(strong("Objective Function and Constraints")),
                  verbatimTextOutput("presc_obj_txt"),
                  h4(strong("Optimal Actions (Controllable Variables)")),
                  tableOutput("optimal_actions"),
                  h4(strong("Predicted Outcome with Optimal Actions")),
                  verbatimTextOutput("predicted_optimal"),
                  h4(strong("Prescriptive Tab Info")),
                  uiOutput("presc_info")
                )
              )
      )
    )
  )
)

# ---- Server ----

server <- function(input, output, session) {
  
  # ---- Home Page ----
  output$data_dictionary <- renderDT(data_dictionary, options = list(pageLength = 15))
  
  # ---- Descriptive Plots ----
  output$histogram <- renderPlot({ plot_histogram })
  output$supplier <- renderPlot({ plot_supplier })
  output$cut_speed <- renderPlot({ plot_cut_speed })
  output$glass_thickness <- renderPlot({ plot_glass_thickness })
  output$window_size <- renderPlot({ plot_window_size })
  
  custom_plot_data <- eventReactive(input$updateGraph, {
    req(input$x_variable, input$y_variable)
    ggplot(data, aes(x = .data[[input$x_variable]], y = .data[[input$y_variable]])) +
      geom_jitter(alpha = 0.5) +
      geom_smooth() +
      theme_classic() +
      theme(
        axis.title.y = element_text(size = 18, margin = margin(r = 10)),
        axis.title.x = element_text(size = 18, margin = margin(t = 10)),
        axis.text = element_text(size = 16)
      )
  }, ignoreNULL = FALSE)
  output$custom_plot <- renderPlot({ custom_plot_data() })
  
  # ---- Predictive Analytics ----
  
  last_fit <- reactiveVal(NULL)
  last_predictors <- reactiveVal(NULL)
  last_response <- reactiveVal(NULL)
  last_method <- reactiveVal(NULL)
  last_coef <- reactiveVal(NULL)
  last_preproc <- reactiveVal(NULL)
  
  observeEvent(input$response, {
    output$model_type_ui <- renderUI({
      if (input$response == "pass_fail") {
        helpText("For glmnet and random forest, logistic regression/classification is used.")
      } else {
        helpText("For glmnet and random forest, regression is used.")
      }
    })
  })
  
  fit_result <- eventReactive(input$fit_model, {
    req(input$predictors)
    preprocess <- input$preprocess
    response <- input$response; predictors <- input$predictors
    if (!"medianImpute" %in% preprocess) {
      df_model <- df[complete.cases(df[, c(response, predictors)]), ]
    } else {
      df_model <- df
    }
    trainIndex <- createDataPartition(df_model[[response]], p = .7, list = FALSE)
    train <- df_model[trainIndex, ]; test <- df_model[-trainIndex, ]
    caret_method <- input$caret_method
    use_cv <- input$use_cv; alpha_val <- ifelse(is.null(input$alpha), 0.5, input$alpha)
    is_class <- (response == "pass_fail")
    if (is_class) {
      train$pass_fail <- factor(train$pass_fail, levels = c("Fail", "Pass"))
      test$pass_fail <- factor(test$pass_fail, levels = c("Fail", "Pass"))
    }
    formula <- as.formula(
      paste0("`", response, "` ~ ", paste0("`", predictors, "`", collapse = " + "))
    )
    tr_ctrl <- trainControl(
      method = if (use_cv) "cv" else "none",
      number = 10,
      classProbs = is_class,
      savePredictions = "final",
      summaryFunction = if (is_class) twoClassSummary else defaultSummary
    )
    pre_proc <- if (length(preprocess) > 0) preprocess else NULL
    if (caret_method == "glmnet") {
      tuneGrid <- expand.grid(alpha = alpha_val, lambda = seq(0.001, 0.1, length = 10))
      fit <- train(
        formula, data = train, method = "glmnet",
        family = if (is_class) "binomial" else "gaussian",
        preProcess = pre_proc, trControl = tr_ctrl, tuneGrid = tuneGrid,
        metric = if (is_class) "ROC" else "RMSE"
      )
      coef_df <- as.data.frame(as.matrix(coef(fit$finalModel, fit$bestTune$lambda)))
      coef_df$term <- rownames(coef_df); rownames(coef_df) <- NULL; names(coef_df)[1] <- "estimate"
    } else if (caret_method == "rf") {
      fit <- train(formula, data = train, method = "rf", preProcess = pre_proc, trControl = tr_ctrl)
      coef_df <- NULL
    } else {
      fit <- train(formula, data = train, method = caret_method, preProcess = pre_proc, trControl = tr_ctrl)
      coef_df <- if (caret_method == "glm") tryCatch(broom::tidy(fit$finalModel), error=function(e) NULL) else NULL
      if (!is.null(coef_df) && "term" %in% names(coef_df)) coef_df$term <- sub("^`|`$", "", coef_df$term)
    }
    if (is_class) {
      pred_train_prob <- predict(fit, train, type = "prob")[, "Pass"]
      pred_test_prob <- predict(fit, test, type = "prob")[, "Pass"]
      pred_train_class <- factor(ifelse(pred_train_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
      pred_test_class <- factor(ifelse(pred_test_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
      y_train <- train[[response]]; y_test <- test[[response]]
      train_metric <- postResample(pred_train_class, y_train)
      test_metric <- postResample(pred_test_class, y_test)
      residuals <- pred_train_prob - as.numeric(y_train == "Pass")
      auc_train <- auc_test <- NA
      suppressWarnings({
        auc_train <- tryCatch({pROC::auc(pROC::roc(as.numeric(y_train), pred_train_prob))}, error = function(e) NA)
        auc_test <- tryCatch({pROC::auc(pROC::roc(as.numeric(y_test), pred_test_prob))}, error = function(e) NA)
      })
    } else {
      pred_train <- predict(fit, train)
      pred_test <- predict(fit, test)
      y_train <- train[[response]]; y_test <- test[[response]]
      train_metric <- postResample(pred_train, y_train)
      test_metric <- postResample(pred_test, y_test)
      auc_train <- auc_test <- NA
      residuals <- y_train - pred_train
    }
    last_fit(fit); last_predictors(predictors); last_response(response)
    last_method(caret_method); last_coef(coef_df); last_preproc(pre_proc)
    list(
      fit = fit, coef = coef_df,
      varimp = tryCatch(varImp(fit)$importance, error=function(e) NULL),
      train_metric = train_metric, test_metric = test_metric,
      auc_train = auc_train, auc_test = auc_test,
      ytrain = y_train, ytest = y_test,
      pred_train = if (is_class) pred_train_prob else pred_train,
      pred_test = if (is_class) pred_test_prob else pred_test,
      residuals = if (!is_class) residuals else NULL,
      model_type = caret_method,
      is_classification = is_class
    )
  })
  
  output$coef_table <- renderDT({
    result <- fit_result(); req(result)
    if (!is.null(result$coef)) {
      coef_tbl <- result$coef
      if ("term" %in% names(coef_tbl)) coef_tbl$term <- gsub("[`\\\\]", "", coef_tbl$term)
      datatable(coef_tbl, rownames = FALSE, options = list(pageLength = 15))
    } else {
      vi <- result$varimp
      dt <- data.frame(term = gsub("[`\\\\]", "", rownames(vi)), importance = vi[,1], row.names = NULL)
      datatable(dt, rownames = FALSE, options = list(pageLength = 15))
    }
  })
  
  output$model_fit_plot <- renderPlot({
    result <- fit_result(); req(result)
    if (!result$is_classification) {
      df_plot <- data.frame(
        Actual = c(result$ytrain, result$ytest),
        Predicted = c(result$pred_train, result$pred_test),
        Data = rep(c("Train", "Test"), c(length(result$ytrain), length(result$ytest)))
      )
      ggplot(df_plot, aes(x = Actual, y = Predicted, color = Data)) +
        geom_point(alpha = 0.6) +
        geom_abline(slope = 1, intercept = 0, linetype = "dashed") +
        theme_minimal() +
        labs(title = "Predicted vs Actual", color = "Data") +
        xlab("Actual") + ylab("Predicted")
    } else {
      df_plot <- data.frame(
        Actual = factor(c(result$ytrain, result$ytest)),
        Predicted = c(result$pred_train, result$pred_test),
        Data = rep(c("Train", "Test"), c(length(result$ytrain), length(result$ytest)))
      )
      ggplot(df_plot, aes(x = Predicted, fill = Actual)) +
        geom_density(alpha = 0.5) +
        facet_wrap(~Data) +
        theme_minimal() +
        labs(title = "Predicted Probability Density", x = "Predicted Probability", fill = "Actual")
    }
  })
  
  # --- Only show residual plot for regression, never show for classification ---
  output$residual_plot_box <- renderUI({
    result <- fit_result()
    if (!is.null(result) && !result$is_classification) {
      tagList(
        h4("Residual Plot"),
        plotOutput("residual_plot")
      )
    } else {
      NULL
    }
  })
  output$residual_plot <- renderPlot({
    result <- fit_result()
    if (!is.null(result) && !result$is_classification) {
      df_resid <- data.frame(Fitted = result$pred_train, Residuals = result$residuals)
      ggplot(df_resid, aes(x = Fitted, y = Residuals)) +
        geom_point(alpha = 0.5) +
        geom_hline(yintercept = 0, linetype = "dashed") +
        theme_minimal() +
        labs(title = "Residuals vs Fitted (Train Set)")
    }
  })
  
  output$fit_summary <- renderPrint({
    result <- fit_result(); req(result)
    cat("Training Set Metrics:\n"); print(result$train_metric)
    cat("\nTest Set Metrics:\n"); print(result$test_metric)
    if (!is.na(result$auc_train)) {
      cat(sprintf("\nAUC (Train): %.3f", result$auc_train))
      cat(sprintf("\nAUC (Test):  %.3f", result$auc_test))
    }
    cat("\n\nModel Details:\n"); print(result$fit)
    if (input$use_cv) {
      cat("\nCross-validation Results (caret):\n"); print(result$fit$results)
    }
    if (input$caret_method == "glmnet") {
      cat("\nBest tuned lambda:\n"); print(result$fit$bestTune)
    }
  })
  
  # ---- Prescriptive Analytics ----
  output$presc_sliders <- renderUI({
    predictors <- last_predictors()
    if (is.null(predictors)) return(tags$p("Please fit a model on the Predictive Analytics tab first.", style = "color:red"))
    lapply(uncontrollable_vars, function(var) {
      if (var %in% predictors && !is.null(df[[var]]) && any(!is.na(df[[var]]))) {
        if (is.numeric(df[[var]]) || is.integer(df[[var]])) {
          rng <- range(df[[var]], na.rm = TRUE)
          sliderInput(
            inputId = paste0("presc_", var),
            label = var,
            min = floor(rng[1]*100)/100,
            max = ceiling(rng[2]*100)/100,
            value = mean(rng, na.rm = TRUE)
          )
        } else if (is.factor(df[[var]]) || is.character(df[[var]])) {
          lvls <- unique(na.omit(df[[var]]))
          selectInput(
            inputId = paste0("presc_", var),
            label = var,
            choices = lvls,
            selected = lvls[1]
          )
        }
      }
    })
  })
  
  output$controllable_bounds_ui <- renderUI({
    predictors <- last_predictors()
    req(predictors)
    controls_in_model <- intersect(controllable_vars, predictors)
    lapply(controls_in_model, function(var) {
      rng <- range(df[[var]], na.rm = TRUE)
      fluidRow(
        column(6, numericInput(paste0("boundmin_", var), paste(var, "min:"), value = rng[1])),
        column(6, numericInput(paste0("boundmax_", var), paste(var, "max:"), value = rng[2]))
      )
    })
  })
  
  presc_unctrl_values <- reactive({
    vals <- sapply(uncontrollable_vars, function(var) {
      input_val <- input[[paste0("presc_", var)]]
      if (is.null(input_val)) NA else input_val
    })
    names(vals) <- uncontrollable_vars
    vals
  })
  
  output$presc_info <- renderUI({
    req(last_fit(), last_response())
    info <- if (last_response() != "breakage_rate" || last_method() != "glm")
      tags$p(style="color:red;", "Please train a linear regression model on 'breakage_rate' (caret method = glm) to use prescriptive analytics. (LPsolve only supports linear regression!)")
    else
      tags$p(style="color:green;", "Model and optimization objective are in sync. (Linear regression only)")
    missing_ctrls <- setdiff(controllable_vars, last_predictors())
    if (length(missing_ctrls) > 0)
      info <- tagList(info, tags$p(style="color:orange;", paste("The following controllable variables are NOT in your predictors and will not be optimized:", paste(missing_ctrls, collapse = ", "))))
    info
  })
  
  map_var <- function(var, coefs_names) {
    clean <- function(x) gsub("[^A-Za-z0-9]", "", x)
    var_clean <- clean(var)
    for (cand in coefs_names) {
      cand_clean <- clean(cand)
      if (identical(tolower(var_clean), tolower(cand_clean))) return(cand)
    }
    NA_character_
  }
  
  output$presc_coef_table <- renderDT({
    fit <- last_fit()
    req(fit)
    coefs <- tryCatch({
      coef_df <- broom::tidy(fit$finalModel)
      if ("term" %in% names(coef_df)) coef_df$term <- gsub("[`\\\\]", "", coef_df$term)
      coef_df
    }, error = function(e) NULL)
    if (!is.null(coefs)) datatable(coefs, rownames = FALSE, options = list(pageLength = 20))
  })
  
  observeEvent(input$presc_optimize, {
    req(last_fit(), last_predictors(), last_response())
    fit <- last_fit()
    predictors <- last_predictors()
    response <- last_response()
    method <- last_method()
    tol <- input$min_br_constraint
    
    controllable_in_model <- intersect(controllable_vars, predictors)
    uncontrollable_in_model <- intersect(uncontrollable_vars, predictors)
    coefs <- coef(fit$finalModel)
    coefs_names <- names(coefs)
    intercept <- unname(coefs[1])
    
    var_order <- character()
    obj <- numeric()
    match_names <- character()
    for (v in controllable_in_model) {
      mapped <- map_var(v, coefs_names)
      if (!is.na(mapped)) {
        var_order <- c(var_order, v)
        obj <- c(obj, coefs[mapped])
        match_names <- c(match_names, mapped)
      } else {
        var_order <- c(var_order, v)
        obj <- c(obj, NA)
        match_names <- c(match_names, NA)
      }
    }
    n_vars <- length(var_order)
    valid_vars <- !is.na(obj)
    if (sum(valid_vars) == 0) {
      output$optimal_actions <- renderTable({data.frame(Variable = controllable_vars, Optimal_Value = NA)})
      output$predicted_optimal <- renderPrint({"Optimization cannot be run: No valid numeric controllable variables."})
      output$presc_obj_txt <- renderPrint({cat("No valid numeric controllable variables in the model after name matching.\n")})
      return()
    }
    var_order <- var_order[valid_vars]
    obj <- obj[valid_vars]
    n_vars <- length(var_order)
    
    bounds <- t(sapply(var_order, function(v) {
      minname <- paste0("boundmin_", v)
      maxname <- paste0("boundmax_", v)
      minval <- input[[minname]]
      maxval <- input[[maxname]]
      if (!is.null(minval) && !is.null(maxval)) {
        c(minval, maxval)
      } else {
        range(df[[v]], na.rm = TRUE)
      }
    }))
    lower <- bounds[,1]
    upper <- bounds[,2]
    
    uncontrollable_coefs <- numeric()
    for (v in uncontrollable_in_model) {
      mapped <- map_var(v, coefs_names)
      uncontrollable_coefs <- c(uncontrollable_coefs, if (!is.na(mapped)) coefs[mapped] else 0)
    }
    unctrl_vals <- presc_unctrl_values()
    unctrl_vals <- unctrl_vals[uncontrollable_in_model]
    for(i in seq_along(unctrl_vals)) {
      if (is.na(unctrl_vals[[i]])) {
        if (is.numeric(df[[names(unctrl_vals)[i]]])) {
          unctrl_vals[[i]] <- mean(df[[names(unctrl_vals)[i]]], na.rm = TRUE)
        } else if (is.factor(df[[names(unctrl_vals)[i]]])) {
          unctrl_vals[[i]] <- levels(df[[names(unctrl_vals)[i]]])[1]
        }
      }
    }
    uncontrollable_num_coefs <- as.numeric(uncontrollable_coefs)
    uncontrollable_num_vals <- suppressWarnings(as.numeric(unctrl_vals))
    uncontrollable_num_coefs[is.na(uncontrollable_num_coefs)] <- 0
    uncontrollable_num_vals[is.na(uncontrollable_num_vals)] <- 0
    
    constr_mat <- matrix(obj, nrow = 1)
    constr_dir <- ">="
    constr_rhs <- tol - intercept - sum(uncontrollable_num_coefs * uncontrollable_num_vals)
    
    bound_mat <- diag(n_vars)
    constr_mat <- rbind(constr_mat,  bound_mat,  bound_mat)
    constr_dir <- c(constr_dir, rep(">=", n_vars), rep("<=", n_vars))
    constr_rhs <- c(constr_rhs, as.numeric(lower), as.numeric(upper))
    
    output$presc_obj_txt <- renderPrint({
      cat("Objective function:\n")
      cat("  Minimize:\n    ")
      cat(paste(sprintf("% .4f × %s", obj, var_order), collapse = "\n    "), "\n")
      cat("Subject to:\n")
      cat("    ")
      cat(paste(sprintf("% .4f × %s", obj, var_order), collapse = " + "))
      cat(sprintf(" >= %.4f\n", constr_rhs[1]))
      for (i in seq_along(var_order)) {
        cat(sprintf("    %s in [%.4f, %.4f]\n", var_order[i], lower[i], upper[i]))
      }
      for (i in seq_along(unctrl_vals)) {
        cat(sprintf("    %s = %s\n", uncontrollable_in_model[i], as.character(unctrl_vals[i])))
      }
      cat(sprintf("    (Intercept) = %.4f\n", intercept))
    })
    
    sol <- tryCatch({
      lpSolve::lp(
        direction = "min",
        objective.in = obj,
        const.mat = constr_mat,
        const.dir = constr_dir,
        const.rhs = constr_rhs,
        all.int = FALSE
      )
    }, error = function(e) NULL)
    
    all_ctrls <- controllable_vars
    opt_ctrl <- rep(NA, length(all_ctrls)); names(opt_ctrl) <- all_ctrls
    if (!is.null(sol) && sol$status == 0) {
      opt_ctrl[var_order] <- sol$solution
      pred <- intercept + sum(obj * sol$solution) + sum(uncontrollable_num_coefs * uncontrollable_num_vals)
      output$optimal_actions <- renderTable({
        data.frame(Variable = all_ctrls, Optimal_Value = round(opt_ctrl, 3))
      })
      output$predicted_optimal <- renderPrint({
        cat("Predicted Minimum Breakage Rate:\n")
        cat(sprintf("    (Constraint: > %.3g)\n    Value: %.6f\n", tol, pred))
        cat("\nOptimal controllable settings are above.\n")
      })
    } else {
      output$optimal_actions <- renderTable({
        data.frame(Variable = all_ctrls, Optimal_Value = NA)
      })
      output$predicted_optimal <- renderPrint({
        cat("Optimization failed:\n")
        cat("    No feasible solution or problem with model coefficients/inputs.\n")
      })
    }
  })
  
}

shinyApp(ui, server)