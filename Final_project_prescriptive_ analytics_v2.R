library(readxl)
library(shiny)
library(caret)
library(DT)
library(ggplot2)
library(broom)

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

# Define controllable and uncontrollable variables
controllable_vars <- c("Cut speed", "Edge Deletion rate", "Spacer Distance", "Silicon Viscosity")
uncontrollable_vars <- c(
  "Ambient Temp", "Window Size", "Glass thickness", "Window color",
  "Window Type", "Glass Supplier", "Glass Supplier Location"
)

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
            choices = c("Breakage Rate", "Pass/Fail"),
            selected = "Breakage Rate"
          ),
          uiOutput("model_type_ui"),
          checkboxGroupInput(
            "predictors",
            "Select Predictors:",
            choices = colnames(df)[!(colnames(df) %in% c("Batch", "Breakage Rate", "Pass/Fail"))],
            selected = colnames(df)[!(colnames(df) %in% c("Batch", "Breakage Rate", "Pass/Fail"))]
          ),
          checkboxGroupInput(
            "preprocess",
            "Preprocessing Steps:",
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
          h4("Model Coefficient Table / Variable Importance"),
          DTOutput("coef_table"),
          h4("Model Fit Evaluation"),
          plotOutput("model_fit_plot"),
          h4("Residual Plot"),
          plotOutput("residual_plot"),
          verbatimTextOutput("fit_summary")
        )
      )
    ),
    tabPanel(
      "Prescriptive Analytics (LPsolve)",
      sidebarLayout(
        sidebarPanel(
          h4("Set Uncontrollable Variables"),
          uiOutput("presc_sliders"),
          h4("Set/Override Bounds for Controllables"),
          uiOutput("controllable_bounds_ui"),
          h4("Optimization Objective"),
          sliderInput("min_br_constraint", "Minimum Predicted Breakage Rate (constraint):",
                      min = 0,
                      max = max(df$`Breakage Rate`, na.rm = TRUE) + 10,
                      value = 0.01,
                      step = 0.01),
          tags$p("Minimize Breakage Rate (constrained above) using linear programming."),
          actionButton("presc_optimize", "Optimize Actions", icon = icon("bolt"))
        ),
        mainPanel(
          h4("Objective Function and Constraints"),
          verbatimTextOutput("presc_obj_txt"),
          h4("Optimal Actions (Controllable Variables)"),
          tableOutput("optimal_actions"),
          h4("Predicted Outcome with Optimal Actions"),
          verbatimTextOutput("predicted_optimal"),
          h4("Prescriptive Tab Info"),
          uiOutput("presc_info")
        )
      )
    )
  )
)

server <- function(input, output, session) {
  last_fit <- reactiveVal(NULL)
  last_predictors <- reactiveVal(NULL)
  last_response <- reactiveVal(NULL)
  last_method <- reactiveVal(NULL)
  last_coef <- reactiveVal(NULL)
  last_preproc <- reactiveVal(NULL)
  
  observeEvent(input$response, {
    output$model_type_ui <- renderUI({
      if (input$response == "Pass/Fail") {
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
    is_class <- (response == "Pass/Fail")
    if (is_class) {
      train$`Pass/Fail` <- factor(train$`Pass/Fail`, levels = c("Fail", "Pass"))
      test$`Pass/Fail` <- factor(test$`Pass/Fail`, levels = c("Fail", "Pass"))
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
  
  output$residual_plot <- renderPlot({
    result <- fit_result(); req(result)
    if (!result$is_classification) {
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
  
  output$presc_sliders <- renderUI({
    predictors <- last_predictors()
    if (is.null(predictors)) return(tags$p("Please train a model first."))
    lapply(uncontrollable_vars, function(var) {
      if (var %in% predictors && !is.null(df[[var]]) && any(!is.na(df[[var]]))) {
        if (is.numeric(df[[var]]) || is.integer(df[[var]])) {
          rng <- range(df[[var]], na.rm = TRUE)
          sliderInput(
            inputId = paste0("presc_", gsub(" ", "_", var)),
            label = var,
            min = floor(rng[1]*100)/100,
            max = ceiling(rng[2]*100)/100,
            value = mean(rng, na.rm = TRUE)
          )
        } else if (is.factor(df[[var]]) || is.character(df[[var]])) {
          lvls <- unique(na.omit(df[[var]])) # remove NA from choices
          selectInput(
            inputId = paste0("presc_", gsub(" ", "_", var)),
            label = var,
            choices = lvls,
            selected = lvls[1]
          )
        }
      }
    })
  })
  
  # UI for user-editable bounds
  output$controllable_bounds_ui <- renderUI({
    predictors <- last_predictors()
    req(predictors)
    controls_in_model <- intersect(controllable_vars, predictors)
    lapply(controls_in_model, function(var) {
      rng <- range(df[[var]], na.rm = TRUE)
      fluidRow(
        column(6, numericInput(paste0("boundmin_", gsub(" ", "_", var)), paste(var, "min:"), value = rng[1])),
        column(6, numericInput(paste0("boundmax_", gsub(" ", "_", var)), paste(var, "max:"), value = rng[2]))
      )
    })
  })
  
  presc_unctrl_values <- reactive({
    vals <- sapply(uncontrollable_vars, function(var) {
      input_val <- input[[paste0("presc_", gsub(" ", "_", var))]]
      if (is.null(input_val)) NA else input_val
    })
    names(vals) <- uncontrollable_vars
    vals
  })
  
  output$presc_info <- renderUI({
    req(last_fit(), last_response())
    info <- if (last_response() != "Breakage Rate" || last_method() != "glm")
      tags$p(style="color:red;", "Please train a linear regression model on 'Breakage Rate' (caret method = glm) to use prescriptive analytics. (LPsolve only supports linear regression!)")
    else
      tags$p(style="color:green;", "Model and optimization objective are in sync. (Linear regression only)")
    missing_ctrls <- setdiff(controllable_vars, last_predictors())
    if (length(missing_ctrls) > 0)
      info <- tagList(info, tags$p(style="color:orange;", paste("The following controllable variables are NOT in your predictors and will not be optimized:", paste(missing_ctrls, collapse = ", "))))
    info
  })
  
  # Robust matching: Remove all non-alphanumeric for comparison
  map_var <- function(var, coefs_names) {
    clean <- function(x) gsub("[^A-Za-z0-9]", "", x)
    var_clean <- clean(var)
    for (cand in coefs_names) {
      cand_clean <- clean(cand)
      if (identical(tolower(var_clean), tolower(cand_clean))) return(cand)
    }
    NA_character_
  }
  
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
    
    # Use user bounds if available
    bounds <- t(sapply(var_order, function(v) {
      minname <- paste0("boundmin_", gsub(" ", "_", v))
      maxname <- paste0("boundmax_", gsub(" ", "_", v))
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