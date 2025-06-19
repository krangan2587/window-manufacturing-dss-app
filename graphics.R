#setwd("C:/Users/micha/Downloads")

library(readxl)
library(tidyverse)

set.seed(27)

df <- read_excel("Window_Manufacturing.xlsx",
                   col_names = c("breakage_rate", "window_size", "glass_thickness",
                                 "ambient_temp", "cut_speed", "edge_deletion_rate",
                                 "spacer_distance", "window_color", "window_type",
                                 "glass_supplier", "silicon_viscosity", "supplier_location"),
                   skip = 1) %>% 
  mutate(pass_fail = factor(ifelse(breakage_rate < 7, "Pass", "Fail")),
         across(where(is.character), as.factor),
         across(where(is.numeric), as.numeric)) %>% 
  select("breakage_rate", sort(colnames(.)))

# Breakage Rate Distribution
breakage_rate_mean <- data.frame(
  column = "breakage_rate",
  mean = round(mean(df$breakage_rate),2)
)

plot_histogram <- df %>% 
  ggplot() +
  geom_histogram(aes(x = breakage_rate),
                 color = "white", fill = "steelblue4") +
  geom_vline(data = breakage_rate_mean,
             aes(xintercept = mean),
             linetype = "dashed", color = "black",
             size = 1,
             show.legend = FALSE) +
  geom_label(data = breakage_rate_mean,
             aes(x = mean + 0.5, y = 200,
                 label = paste("Mean:", mean)),
             hjust = 0) +
  scale_y_continuous(expand = c(0,0)) +
  scale_x_continuous(breaks = seq(from = 0, to = 40, by = 5)) +
  labs(y = "Frequency",
       x = "Breakage Rate") +
  theme_classic() +
  theme(
    axis.title.y = element_text(size = 18,
                                margin = margin(r = 10)),
    axis.title.x = element_text(size = 18,
                                margin = margin(t = 10)),
    axis.text = element_text(size = 16)
)

# Breakage Rate by Glass Supplier
supplier_mean <- subset(df, !is.na(glass_supplier)) %>% 
  group_by(glass_supplier) %>% 
  summarise(mean = round(mean(breakage_rate),2)) %>% 
  mutate(annotate = mean + 2)

plot_supplier <- subset(df, !is.na(glass_supplier)) %>% 
  ggplot(aes(x = glass_supplier,
             y = breakage_rate)) +
  geom_jitter(aes(color = glass_supplier),
              show.legend = FALSE,
              alpha = 0.5, size = 2) +
  stat_summary(fun = mean, geom = "crossbar", color = "black", size = 0.5) +
  geom_label(data = supplier_mean, 
            aes(x = glass_supplier, y = annotate, 
                label = paste("Mean:",mean),
                size = 12, fontface = "bold",
                fill = glass_supplier,
                alpha = 0.5),
            show.legend = FALSE) +
  labs(title = NULL,
       x = NULL,
       y = "Breakage Rate") +
  theme_classic() +
  theme(
    axis.title.y = element_text(size = 18, 
                                margin = margin(r = 10)),
    axis.text.x = element_text(size = 18, color = "black", 
                               margin = margin(t = 5, b = 5)),
    axis.text = element_text(size = 16)
)

# Breakage Rate by Cut Speed
plot_cut_speed <- df %>% 
  ggplot(aes(x = cut_speed, 
             y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 2) +
  geom_smooth() +
  scale_x_continuous(breaks = seq(from = 0.5, to = 3, by = 0.5)) +
  labs(x = "Cut Speed (meters per minute)",
       y = "Breakage Rate") +
  theme_classic() +
  theme(
    axis.title.y = element_text(size = 18,
                                margin = margin(r = 10)),
    axis.title.x = element_text(size = 18,
                                margin = margin(t = 10)),
    axis.text = element_text(size = 16)
)

# Breakage Rate by Glass Thickness
plot_glass_thickness <- df %>% 
  ggplot(aes(x = glass_thickness, 
             y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 1.75) +
  geom_smooth() +
  labs(x = "Glass Thickness (inches)",
       y = "Breakage Rate") +
  theme_classic() +
  theme(
    axis.title.y = element_text(size = 18,
                                margin = margin(r = 10)),
    axis.title.x = element_text(size = 18,
                                margin = margin(t = 10)),
    axis.text = element_text(size = 16)
)

# Breakage Rate by Window Size
plot_window_size <- df %>% 
  ggplot(aes(x = window_size, 
             y = breakage_rate)) +
  geom_point(alpha = 0.5, size = 1.75) +
  geom_smooth() +
  scale_x_continuous(breaks = seq(from = 50, to = 80, by = 5)) +
  labs(x = "Diagonal Window Size (inches)",
       y = "Breakage Rate") +
  theme_classic() +
  theme(
    axis.title.y = element_text(size = 18,
                                margin = margin(r = 10)),
    axis.title.x = element_text(size = 18,
                                margin = margin(t = 10)),
    axis.text = element_text(size = 16)
)

