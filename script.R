# set working directory
setwd("your/working/directory")

# import libraries
library(readxl)
library(tidyverse)
library(gtsummary)
library(flextable)
library(ggpubr)

# import data
df <- read_excel("your/working/directory/data.xlsx")
attach(df)
df %>% colnames

# create a summary statistics table
table1 <- df[,-1] %>% # remove the first variable (ID)
            tbl_summary(by = "Retinopathy",
                        type = list(where(is.numeric) ~ "continuous"),
                        missing_text = "Missing Data") %>%
            add_p() %>%
            add_overall()

# save the table as a docx file
table1 %>%
  as_flex_table() %>%
  save_as_docx(path = "Table 1.docx")

# univariate logistic regression
table2 <- df[,c(17,2,3,18,4,5,7,10,11,14)] %>%
            tbl_uvregression(method = glm, 
                             method.args = list(family = "binomial"),
                             y = "Retinopathy",
                             exponentiate = TRUE) # get odds ratios

table2 %>%
  as_flex_table() %>%
  save_as_docx(path = "Table 2.docx")

# plot probability curves
# build the model
mod1 <- glm(Retinopathy ~ `Sex`, family = "binomial")

# create prediction data frame
preds <- predict(mod1, se.fit = TRUE, type = "response")

# create confidence intervals
pred1 <- preds$fit
upr1 <- preds$fit + 1.96 * preds$se.fit
lwr1 <- preds$fit - 1.96 * preds$se.fit

# probability plots for each independent variable
plot1 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`Age (years)`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot1,
       filename = "Plot 1.png",
       height = 4,
       width = 8,
       dpi = 900)

df_summary <- df %>%
  group_by(Sex) %>%
  summarise(prob = mean(Retinopathy == "1"),
            n = n(),
            se = sqrt(prob * (1 - prob) / n),
            lower_ci = prob - 1.96 * se,
            upper_ci = prob + 1.96 * se)
plot2 <- ggplot(df_summary, aes(x = Sex, y = prob)) +
           geom_point(stat = "identity", position = "dodge", size = 3) +
           geom_errorbar(aes(ymin = lower_ci, ymax = upper_ci), width = 0.2, position = position_dodge(width = 0.9)) +
           labs(y = "Probability of Retinopathy")
ggsave(plot2,
       filename = "Plot 2.png",
       height = 4,
       width = 8,
       dpi = 900)

plot3 <- df %>%
          mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
          ggplot(aes(BMI, prob)) +
          geom_point(alpha = 0.2) +
          geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
          labs(y = "Probability of Retinopathy")
ggsave(plot3,
       filename = "Plot 3.png",
       height = 4,
       width = 8,
       dpi = 900)

plot4 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`BSL Fasting`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot4,
       filename = "Plot 4.png",
       height = 4,
       width = 8,
       dpi = 900)

plot5 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`BSL PP`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot5,
       filename = "Plot 5.png",
       height = 4,
       width = 8,
       dpi = 900)

plot6 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`HbA1c`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot6,
       filename = "Plot 6.png",
       height = 4,
       width = 8,
       dpi = 900)

plot7 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`Duration of DM (years)`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot7,
       filename = "Plot 7.png",
       height = 4,
       width = 8,
       dpi = 900)

df_summary <- df %>%
  group_by(Metformin) %>%
  summarise(prob = mean(Retinopathy == "1"),
            n = n(),
            se = sqrt(prob * (1 - prob) / n),
            lower_ci = prob - 1.96 * se,
            upper_ci = prob + 1.96 * se)
plot8 <- ggplot(df_summary, aes(x = Metformin, y = prob)) +
  geom_point(stat = "identity", position = "dodge", size = 3) +
  geom_errorbar(aes(ymin = lower_ci, ymax = upper_ci), width = 0.2, position = position_dodge(width = 0.9)) +
  labs(y = "Probability of Retinopathy")
ggsave(plot8,
       filename = "Plot 8.png",
       height = 4,
       width = 8,
       dpi = 900)

plot9 <- df %>%
  mutate(prob = ifelse(Retinopathy == "1", 1, 0)) %>%
  ggplot(aes(`Number of anti-DM Medications`, prob)) +
  geom_point(alpha = 0.2) +
  geom_smooth(method = "glm", method.args = list(family = "binomial"), color = "black") +
  labs(y = "Probability of Retinopathy")
ggsave(plot9,
       filename = "Plot 9.png",
       height = 4,
       width = 8,
       dpi = 900)
