#install.packages("readxl")
library(readxl)
library(shiny)
library(caret)
library(DT)
library(ggplot2)
library(broom)

# Simulate some example data (replace with your real data reading step)
set.seed(42)
#read the xlsx file
df <- read_excel("Window_Manufacturing.xlsx")
df$`Pass/Fail` <- factor(ifelse(df$`Breakage Rate` < 7, "Pass", "Fail"))

# Ensure all columns are in the correct format
df <- df %>%
  mutate(across(where(is.character), as.factor)) %>%
  mutate(across(where(is.numeric), as.numeric))

# Remove any accidental leading/trailing spaces from column names
colnames(df) <- trimws(colnames(df))

ui <- fluidPage(
  titlePanel("Window Manufacturing DSS Prototype"),
  tabsetPanel(
    tabPanel(
      "Predictive Analytics",
      sidebarLayout(
        sidebarPanel(
          h4("Model Settings"),
          selectInput(
            "response",
            "Choose Response Variable:",
            choices = c("Breakage Rate","Pass/Fail"),
            selected = "Breakage Rate"
          ),
          uiOutput("model_type_ui"),
          checkboxGroupInput(
            "predictors",
            "Select Predictors:",
            choices = colnames(df)[!(colnames(df) %in% c("Batch", "Breakage Rate", "Yield", "YldFrctn", "Pass/Fail"))],
            selected = c("Window Type", "Glass Thickness", "Glass Supplier", "Ambient Temp")
          ),
          checkboxGroupInput(
            "preprocess",
            "Preprocessing Steps:",
            choices = c(
              "Center" = "center",
              "Scale" = "scale",
              "Impute Missing (median)" = "medianImpute",
              "Remove Near Zero Var" = "nzv"
            ),
            selected = c("center", "scale")
          ),
          checkboxInput("use_cv", "10-Fold Cross-Validation", value = TRUE),
          selectInput("caret_method", "Model Type (caret):",
                      choices = c(
                        "Linear/Logistic Regression" = "glm", 
                        "Elastic Net (glmnet)" = "glmnet"
                      ),
                      selected = "glm"
          ),
          conditionalPanel(
            condition = "input.caret_method == 'glmnet'",
            sliderInput("alpha", "Elastic Net Mixing (alpha):", min = 0, max = 1, value = 0.5)
          ),
          actionButton("fit_model", "Fit Model", icon = icon("cogs"))
        ),
        mainPanel(
          h4("Model Coefficient Table"),
          DTOutput("coef_table"),
          h4("Model Fit Evaluation"),
          plotOutput("model_fit_plot"),
          verbatimTextOutput("fit_summary")
        )
      )
    )
  )
)

server <- function(input, output, session) {
  observeEvent(input$response, {
    output$model_type_ui <- renderUI({
      if (input$response %in% c("Pass/Fail")) {
        helpText("For glmnet, logistic regression is used.")
      } else {
        helpText("For glmnet, elastic net regression is used.")
      }
    })
  })
  
  data_split <- reactive({
    req(input$predictors)
    df_model <- df
    # Remove rows with NA in predictors or response
    df_model <- df_model[complete.cases(df_model[, c(input$response, input$predictors)]), ]
    trainIndex <- createDataPartition(df_model[[input$response]], p = .7, list = FALSE)
    list(train = df_model[trainIndex, ], test = df_model[-trainIndex, ])
  })
  
  fit_result <- eventReactive(input$fit_model, {
    req(input$predictors)
    split <- data_split()
    train <- split$train
    test <- split$test
    response <- input$response
    predictors <- input$predictors
    preprocess <- input$preprocess
    caret_method <- input$caret_method
    use_cv <- input$use_cv
    alpha_val <- ifelse(is.null(input$alpha), 0.5, input$alpha)
    # Make sure Pass/Fail is always factor with levels
    if ("Pass/Fail" %in% response) {
      train$`Pass/Fail` <- factor(train$`Pass/Fail`, levels = c("Fail", "Pass"))
      test$`Pass/Fail` <- factor(test$`Pass/Fail`, levels = c("Fail", "Pass"))
    }
    formula <- as.formula(
      paste0("`", response, "` ~ ", paste0("`", predictors, "`", collapse = " + "))
    )
    # Set up training control
    tr_ctrl <- trainControl(
      method = if (use_cv) "cv" else "none",
      number = 10,
      classProbs = response == "Pass/Fail",
      savePredictions = "final"
    )
    # Preprocessing
    pre_proc <- if (length(preprocess) > 0) preprocess else NULL
    
    # For glmnet, we must one-hot encode factors so use dummyVars
    if (caret_method == "glmnet") {
      tuneGrid <- expand.grid(alpha = alpha_val, lambda = seq(0.001, 0.1, length = 10))
      if (response == "Pass/Fail") {
        fit <- train(
          formula, data = train,
          method = "glmnet",
          family = "binomial",
          preProcess = pre_proc,
          trControl = tr_ctrl,
          tuneGrid = tuneGrid
        )
        pred_train_prob <- predict(fit, train, type = "prob")[, "Pass"]
        pred_test_prob <- predict(fit, test, type = "prob")[, "Pass"]
        pred_train <- factor(ifelse(pred_train_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
        pred_test <- factor(ifelse(pred_test_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
        y_train <- train[[response]]
        y_test <- test[[response]]
        train_metric <- postResample(pred_train, y_train)
        test_metric <- postResample(pred_test, y_test)
        list(
          fit = fit,
          coef = as.data.frame(as.matrix(coef(fit$finalModel, fit$bestTune$lambda))),
          train_metric = train_metric,
          test_metric = test_metric,
          ytrain = y_train,
          ytest = y_test,
          pred_train = pred_train_prob,
          pred_test = pred_test_prob,
          model_type = "glmnet"
        )
      } else { # regression
        fit <- train(
          formula, data = train,
          method = "glmnet",
          family = "gaussian",
          preProcess = pre_proc,
          trControl = tr_ctrl,
          tuneGrid = tuneGrid
        )
        pred_train <- predict(fit, train)
        pred_test <- predict(fit, test)
        y_train <- train[[response]]
        y_test <- test[[response]]
        train_metric <- postResample(pred_train, y_train)
        test_metric <- postResample(pred_test, y_test)
        list(
          fit = fit,
          coef = as.data.frame(as.matrix(coef(fit$finalModel, fit$bestTune$lambda))),
          train_metric = train_metric,
          test_metric = test_metric,
          ytrain = y_train,
          ytest = y_test,
          pred_train = pred_train,
          pred_test = pred_test,
          model_type = "glmnet"
        )
      }
    } else {
      # glm (caret handles both regression and binary classification)
      if (response == "Pass/Fail") {
        fit <- train(formula, data = train, method = "glm", family = binomial,
                     preProcess = pre_proc, trControl = tr_ctrl)
        pred_train_prob <- predict(fit, train, type = "prob")[, "Pass"]
        pred_test_prob <- predict(fit, test, type = "prob")[, "Pass"]
        pred_train <- factor(ifelse(pred_train_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
        pred_test <- factor(ifelse(pred_test_prob > 0.5, "Pass", "Fail"), levels = c("Fail", "Pass"))
        y_train <- train[[response]]
        y_test <- test[[response]]
        train_metric <- postResample(pred_train, y_train)
        test_metric <- postResample(pred_test, y_test)
        list(
          fit = fit,
          coef = tidy(fit$finalModel),
          train_metric = train_metric,
          test_metric = test_metric,
          ytrain = y_train,
          ytest = y_test,
          pred_train = pred_train_prob,
          pred_test = pred_test_prob,
          model_type = "glm"
        )
      } else {
        fit <- train(formula, data = train, method = "glm", 
                     preProcess = pre_proc, trControl = tr_ctrl)
        pred_train <- predict(fit, train)
        pred_test <- predict(fit, test)
        y_train <- train[[response]]
        y_test <- test[[response]]
        train_metric <- postResample(pred_train, y_train)
        test_metric <- postResample(pred_test, y_test)
        list(
          fit = fit,
          coef = tidy(fit$finalModel),
          train_metric = train_metric,
          test_metric = test_metric,
          ytrain = y_train,
          ytest = y_test,
          pred_train = pred_train,
          pred_test = pred_test,
          model_type = "glm"
        )
      }
    }
  })
  
  output$coef_table <- renderDT({
    result <- fit_result()
    req(result)
    coef_tbl <- result$coef
    # For glmnet, rownames are variable names
    if (result$model_type == "glmnet") {
      coef_tbl <- data.frame(term = rownames(coef_tbl), estimate = coef_tbl[,1], row.names = NULL)
      datatable(
        coef_tbl,
        rownames = FALSE,
        options = list(pageLength = 15)
      )
    } else {
      datatable(
        coef_tbl[, c("term", "estimate", "std.error", "statistic", "p.value")],
        rownames = FALSE,
        options = list(pageLength = 15)
      )
    }
  })
  
  output$model_fit_plot <- renderPlot({
    result <- fit_result()
    req(result)
    is_classification <- input$response == "Pass/Fail"
    if (!is_classification) {
      # Scatter plot: Predicted vs Actual
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
      # Show predicted probabilities for classification
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
  
  output$fit_summary <- renderPrint({
    result <- fit_result()
    req(result)
    cat("Training Set Metrics:\n")
    print(result$train_metric)
    cat("\nTest Set Metrics:\n")
    print(result$test_metric)
    cat("\n\nModel Details:\n")
    print(result$fit)
    if (input$use_cv) {
      cat("\nCross-validation Results (caret):\n")
      print(result$fit$results)
    }
    if (input$caret_method == "glmnet") {
      cat("\nBest tuned lambda:\n")
      print(result$fit$bestTune)
    }
  })
}

shinyApp(ui, server)