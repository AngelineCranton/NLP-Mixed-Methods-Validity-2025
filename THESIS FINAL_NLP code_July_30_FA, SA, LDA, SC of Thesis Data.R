###################### IMPORTS ######################

# Load necessary libraries
library(dplyr)     # For data manipulation and plotting
library(stringr)   # For text cleaning
library(readr)     # For reading CSV data files
library(ggplot2)   # For plotting
library(wordcloud) # For creating the word cloud
library(RColorBrewer) # For creating the word cloud
library(png)       # For saving the word cloud as a png
library(syuzhet)   # For sentiment analysis
library(openxlsx)  # For excel exporting
library(topicmodels) # For LDA topic modeling (Latent Dirichlet Allocation)
library(tm) # For text mining - required for LDA and NRC clustering
library(cluster) # For clustering - required for NRC clustering

# When importing the CSV file, ensure to locate the pop-up box to select the document where it is saved on your computer - a more portable approach than using a specified file path
# The imported data file was previously converted to UTF-CSV for running efficiency
# Select CSV data file - select via file explorer dialogue box (it will pop up on your desktop)
csv_file_path <- file.choose()

###################### GLOBAL VARIABLES ######################

# Read CSV data
df <- read_csv(csv_file_path, col_types = cols())

# Define consolidated terms for mapping - defined by researcher using terminology from the course and participant reflection responses
# Alternative to researcher-defined terms = Latent Dirichlet Allocation (LDA) which derives relevant terminology from the data based on frequency.
consolidated_terms <- list(
  "acquiring grit" = c("acquiring grit", "grit", "courage", "drive", "driven"),
  "help-seeking" = c("asking for help", "help seeking", "seeking help"),
  "backwards scheduling" = c("backwards scheduling", "schedules", "schedule", "scheduling", "time management"),
  "burn out" = c("burn out", "burnt out", "exhausted", "exhaustion", "exhausting", "fatigue", "fatigued", "tired", "tiring", "overworked", "overworking", "depleting", "depleted"),
  "coping with setbacks" = c("setback", "setbacks", "failed", "failing", "failure", "academic setbacks", "growth", "growth mindset", "resilience", "resilient"),
  "cornell note-taking" = c("cornell note", "cornell note-taking", "cornell notes"),
  "curbing procrastination" = c("curb procrastination", "curbing procrastination", "procrastination", "procrastinating", "procrastinate"),
  "goal-setting" = c("goal setting", "setting goals", "goals", "making goals", "creating goals"),
  "learning" = c("learning", "learned", "understand", "understanding", "comprehend", "comprehending", "comprehension", "deep learning", "retention", "memory"),
  "motivation" = c("motivation", "motivate", "motivates", "motivated", "motivating", "carrots", "carrot", "motivational factors", "rewards", "reward", "rewarding"),
  "sleep hygiene" = c("sleep hygiene", "sleep", "sleeping"),
  "stress management" = c("stress", "stress management", "stress as load", "as load", "and load", "course load", "overwhelmed", "overwhelm", "stress as worry", "as worry", "worried", "worrying", "and worry", "anxious", "anxiety", "negative thinking", "negative thoughts", "rumination", "ruminate", "ruminative thinking", "spiral", "spiralling", "catastrophize", "catastrophizing", "relaxation breathing", "breathing", "mindfulness", "fact-checking", "fact checking"),
  "study habits" = c("study", "studying", "study", "study strategies", "study skills"),
  "study skills" = c("pomodoro technique", "pomodoros", "breaks", "taking breaks", "spaced repetition", "interleaved practice", "deliberate practice", "active recall", "cue cards", "flashcards", "flash cards", "que cards"),
  "to-do lists" = c("to-do lists", "to do list", "to-do list", "to do lists", "todo list", "todo lists"),
  "well-being" = c("well-being", "well being", "happy", "happiness", "satisfaction", "satisfied", "satisfying", "perma", "PERMA", "perma theory", "positive emotions", "engagement", "relationships", "meaning", "achievement", "mental health", "self-esteem", "depression", "depressed", "mental illness", "mental health condition"),
  "course satisfaction" = c("satisfaction", "satisfied", "satisfying", "enjoyed", "enjoy", "positive", "liked", "appreciated", "appreciate", "positive experience"),
  "course helpfulness" = c("helpful", "helped", "helps", "useful", "benefitted", "benefit", "beneficial", "valued", "value", "valuable", "practical", "practicality", "applicable", "application", "effective", "works", "worked", "impactful", "impact", "support", "supportive", "supporting")
)

############# DEFINING FUNCTIONS - NLP ANALYSES OF CLEANED TEXT #############

#### TEXT CLEANING ####
# Function to clean the text
Clean_String <- function(string) {
  if (is.na(string) || string == "") {
    print("text is not cleaned. returned empty")
    return("")
  }
  temp <- stringr::str_replace_all(string, "[^a-zA-Z\\s'.-]", " ")
  temp <- stringr::str_replace_all(temp, "[\\s]+", " ")
  temp <- tolower(temp)
  return(temp)
}

#### FREQUENCY ANALYSIS ####
# Function to perform frequency analysis of target strings (consolidated terms)
Count_Targets <- function(cleaned_text, consolidated_terms) {
  target_count <- numeric(length(consolidated_terms))
  for (i in seq_along(consolidated_terms)) {
    consolidated_target <- consolidated_terms[[i]]
    
    if (is.character(consolidated_target)) {
      # Single term mapping
      target_count[i] <- sum(stringr::str_count(cleaned_text, regex(paste0("\\b", consolidated_target, "\\b"), ignore_case = TRUE)))
    } else {
      # Multiple terms mapping
      target_count[i] <- sum(sapply(consolidated_target, function(t) sum(stringr::str_count(cleaned_text, regex(paste0("\\b", t, "\\b"), ignore_case = TRUE)))))
    }
  }
  
  return(data.frame(Target_String = names(consolidated_terms), Present_Count = target_count))
}

#### SENTIMENT ANALYSIS (NRC LEXICON) ####
# Function to perform sentiment analysis
Get_Sentiment <- function(cleaned_text) {
  sentiment <- get_nrc_sentiment(cleaned_text)
  return(sentiment)
}

#### SENTIMENT SHIFT ANALYSIS (NRC LEXICON) ####
# Function to perform sentiment shift analysis between 2 questions
calculate_sentiment_shift <- function(sentiment_df1, sentiment_df2) {
  common_ids <- intersect(sentiment_df1$anonymous_id, sentiment_df2$anonymous_id)
  shift_df <- data.frame(anonymous_id = common_ids)
  for (emotion in colnames(sentiment_df1)[-1]) {
    shift_df[[emotion]] <- sentiment_df2[sentiment_df2$anonymous_id %in% common_ids, emotion] - 
      sentiment_df1[sentiment_df1$anonymous_id %in% common_ids, emotion]
  }
  return(shift_df)
}

#### NRC CLUSTERING (NRC LEXICON) ####
# Function to perform NRC clustering on aggregated sentiments
Perform_NRC_Clustering <- function(aggregated_sentiment_df, num_clusters = 5) {
  # Extract the sentiment scores (exclude the 'Sentiment' column)
  sentiment_matrix <- as.matrix(aggregated_sentiment_df[, -1, drop = FALSE])
  
  # Apply k-means clustering to the sentiment scores
  set.seed(1234)  # For reproducibility
  kmeans_result <- kmeans(sentiment_matrix, centers = num_clusters)
  
  # Add the cluster assignment to the aggregated_sentiment_df
  aggregated_sentiment_df$Cluster <- kmeans_result$cluster
  
  return(aggregated_sentiment_df)
}

#### LATENT DIRICHLET ALLOCATION (LDA) TOPIC MODELING ####
# Function to perform LDA topic modeling with automated stop-word removal
Perform_LDA <- function(text_corpus, num_topics = 18, top_terms = 5) {
  corpus <- Corpus(VectorSource(text_corpus))
  corpus <- tm_map(corpus, content_transformer(tolower))
  corpus <- tm_map(corpus, removePunctuation)
  corpus <- tm_map(corpus, removeNumbers)
  corpus <- tm_map(corpus, removeWords, stopwords("en"))
  corpus <- tm_map(corpus, stripWhitespace)
  
  # Create document-term matrix (DTM)
  dtm <- DocumentTermMatrix(corpus, control = list(wordLengths = c(3, 15)))
  
  # Remove empty rows (documents with no terms)
  rowTotals <- apply(dtm, 1, sum)  
  dtm <- dtm[rowTotals > 0, ]      
  
  # Check if the DTM is empty after filtering
  if (nrow(dtm) == 0) {
    warning("No valid documents for LDA; All documents were empty after preprocessing.")
    return(NULL)
  }
  
  # To perform LDA topic modeling
  lda_model <- LDA(dtm, k = num_topics, control = list(seed = 1234))
  
  # Retrieve the top terms for each topic
  lda_terms <- terms(lda_model, top_terms)
  
  # Print the top terms for each topic
  print("Top terms for each topic:")
  print(lda_terms)
  
  return(lda_terms)
}


###################### MAIN CODE ######################
# CONDUCTING FREQUENCY, SENTIMENT, NRC CLUSTERING, AND LDA ANALYSES ####

# Assign anonymous IDs to each participant
df <- df %>% mutate(anonymous_id = paste0("ID_", 1:n()))

# Initialize lists to store results for Parts 1, 2, 3, 4, 8
count_tables_list <- list(
  P1 = list(), 
  P2_1 = list(), P2_2 = list(), 
  P3_1 = list(), P3_2 = list(), P3_3 = list(), P3_4 = list(), P3_5 = list(), 
  P4_1 = list(), P4_2 = list(), 
  P8_1 = list(), P8_2 = list()
)

sentiment_results <- list(
  P1 = data.frame(), 
  P2_1 = data.frame(), P2_2 = data.frame(), 
  P3_1 = data.frame(), P3_2 = data.frame(), P3_3 = data.frame(), P3_4 = data.frame(), P3_5 = data.frame(), 
  P4_1 = data.frame(), P4_2 = data.frame(), 
  P8_1 = data.frame(), P8_2 = data.frame()
)

nrc_cluster_results <- list(
  P1 = list(), 
  P2_1 = list(), P2_2 = list(), 
  P3_1 = list(), P3_2 = list(), P3_3 = list(), P3_4 = list(), P3_5 = list(), 
  P4_1 = list(), P4_2 = list(), 
  P8_1 = list(), P8_2 = list()
)

lda_results <- list(
  P1 = list(), 
  P2_1 = list(), P2_2 = list(), 
  P3_1 = list(), P3_2 = list(), P3_3 = list(), P3_4 = list(), P3_5 = list(), 
  P4_1 = list(), P4_2 = list(), 
  P8_1 = list(), P8_2 = list()
)

#Define column mappings for Parts 1, 2, 3, 4, 8
question_columns <- list(
  P1 = 5:9,    # Part 1, Combining subsections 1, 2, 3, 4, 5
  P2_1 = 11,   # Part 2, Subsection 1
  P2_2 = 12,   # Part 2, Subsection 2
  P3_1 = 13,   # Part 3, Subsection 1
  P3_2 = 16,   # Part 3, Subsection 2
  P3_3 = 19,   # Part 3, Subsection 3
  P3_4 = 22,   # Part 3, Subsection 4
  P3_5 = 25,   # Part 3, Subsection 5
  P4_1 = 28,   # Part 4, Subsection 1
  P4_2 = 29,   # Part 4, Subsection 2
  P8_1 = 37,   # Part 8, Subsection 1
  P8_2 = 38    # Part 8, Subsection 2
)

############# DATA PROCESSING #############
# Loop through each participant by row to process data
for (i in 1:nrow(df)) {
  # Assign the participant's anonymous ID
  participant_id <- df$anonymous_id[i]
  
  # Process Parts 1, 2, 3, 4, and 8 using column mappings
  for (part_name in names(question_columns)) {
    part_columns <- question_columns[[part_name]]  # Get the column range for the current part
    
    # Extract text from the relevant columns
    cell_text <- paste(df[i, part_columns], collapse = " ")  # Combine multiple columns if necessary
    
    # Skip empty cells and missing values - while retaining the participant's data
    if (is.na(cell_text) || cell_text == "" || grepl("^\\s*$", cell_text)) {
      cat("Empty response for Participant ID:", participant_id, "- Storing empty results.\n")
      
      # Store empty results to ensure every participant is included
      count_tables_list[[part_name]][[participant_id]] <- data.frame(anonymous_id = participant_id, Target_String = NA, Present_Count = NA)
      sentiment_results[[part_name]] <- rbind(sentiment_results[[part_name]], data.frame(anonymous_id = participant_id, sentiment = NA))
      lda_results[[part_name]][[participant_id]] <- list()
      next
    }
    
    # Clean the text, perform frequency and sentiment analysis
    cleaned_text <- Clean_String(cell_text)
    count_table <- Count_Targets(cleaned_text, consolidated_terms)
    sentiment <- Get_Sentiment(cleaned_text)
    lda_terms <- Perform_LDA(cleaned_text)
    
    # Store results dynamically for each part
    count_table <- count_table %>% mutate(anonymous_id = participant_id)  # Add anonymous_id to the count_table
    count_tables_list[[part_name]][[participant_id]] <- count_table
    sentiment_results[[part_name]] <- rbind(sentiment_results[[part_name]], cbind(data.frame(anonymous_id = participant_id), sentiment))
    lda_results[[part_name]][[participant_id]] <- lda_terms
  }
  
  # Debugging: Print results for each participant
  cat("Participant ID:", participant_id, "\n")
  for (part_name in names(question_columns)) {
    if (!is.null(count_tables_list[[part_name]][[participant_id]])) {
      print(count_tables_list[[part_name]][[participant_id]])
      print(sentiment_results[[part_name]][nrow(sentiment_results[[part_name]]), ])
      print(lda_results[[part_name]][[participant_id]])
    }
  }
}

####### SENTIMENT SHIFT ANALYSIS ####### 
# Calculate sentiment shift between the following pairs: P2_1 & P2_2, P8_1 & P8_2
sentiment_shift_P2 <- calculate_sentiment_shift(sentiment_results$P2_1, sentiment_results$P2_2)
sentiment_shift_P8 <- calculate_sentiment_shift(sentiment_results$P8_1, sentiment_results$P8_2)

# Print sentiment shift for both pairs
print(sentiment_shift_P2)
print(sentiment_shift_P8)

####### DATA AGGREGATION FOR NRC CLUSTERING & PLOTTING #######
# Aggregation function for frequency counts - initialize empty data frame to store aggregated counts, summarize counts by target string, and reorder frequency and sentiment scores in descending order per question. 
aggregate_counts <- function(count_tables) {
  aggregated_counts <- data.frame(Target_String = character(), Present_Count = numeric(), stringsAsFactors = FALSE)
  for (count_table in count_tables) {
    aggregated_counts <- rbind(aggregated_counts, count_table)
  }
  aggregated_counts <- aggregated_counts %>%
    group_by(Target_String) %>%
    summarise(Present_Count = sum(Present_Count)) %>%
    arrange(desc(Present_Count))
  return(aggregated_counts)
}

# Aggregation function for sentiment counts
aggregate_sentiment <- function(sentiment_results) {
  sentiment_summary <- colSums(sentiment_results[, -1])  # Exclude 'anonymous_id' column
  sentiment_df <- data.frame(Sentiment = names(sentiment_summary), Score = sentiment_summary)
  sentiment_df <- sentiment_df %>% arrange(desc(Score))
  return(sentiment_df)
}

# Initialize lists for aggregated frequency and sentiment results
aggregated_counts <- list(
  P1 = aggregate_counts(count_tables_list$P1),
  P2_1 = aggregate_counts(count_tables_list$P2_1),
  P2_2 = aggregate_counts(count_tables_list$P2_2),
  P3_1 = aggregate_counts(count_tables_list$P3_1),
  P3_2 = aggregate_counts(count_tables_list$P3_2),
  P3_3 = aggregate_counts(count_tables_list$P3_3),
  P3_4 = aggregate_counts(count_tables_list$P3_4),
  P3_5 = aggregate_counts(count_tables_list$P3_5),
  P4_1 = aggregate_counts(count_tables_list$P4_1),
  P4_2 = aggregate_counts(count_tables_list$P4_2),
  P8_1 = aggregate_counts(count_tables_list$P8_1),
  P8_2 = aggregate_counts(count_tables_list$P8_2)
)

aggregated_sentiments <- list(
  P1 = aggregate_sentiment(sentiment_results$P1),
  P2_1 = aggregate_sentiment(sentiment_results$P2_1),
  P2_2 = aggregate_sentiment(sentiment_results$P2_2),
  P3_1 = aggregate_sentiment(sentiment_results$P3_1),
  P3_2 = aggregate_sentiment(sentiment_results$P3_2),
  P3_3 = aggregate_sentiment(sentiment_results$P3_3),
  P3_4 = aggregate_sentiment(sentiment_results$P3_4),
  P3_5 = aggregate_sentiment(sentiment_results$P3_5),
  P4_1 = aggregate_sentiment(sentiment_results$P4_1),
  P4_2 = aggregate_sentiment(sentiment_results$P4_2),
  P8_1 = aggregate_sentiment(sentiment_results$P8_1),
  P8_2 = aggregate_sentiment(sentiment_results$P8_2)
)

#### PERFORMING NRC CLUSTERING ON AGGREGATED_SENTIMENTS ####
# Perform NRC clustering on aggregated sentiments for each part
nrc_cluster_results <- list()
for (part_name in names(aggregated_sentiments)) {
  nrc_cluster_results[[part_name]] <- Perform_NRC_Clustering(aggregated_sentiments[[part_name]])
}

# Print aggregated and NRC clustering results for verification
print(aggregated_counts)
print(aggregated_sentiments)
print(nrc_cluster_results)

#### PLOTTING/VISUALIZING AGGREGATED FREQUENCY AND SENTIMENT RESULTS ####
## Combining P3 sub-part aggregation results for plotting
# Define P3 sub-parts
p3_subparts <- c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")

# Combine frequency counts for P3 sub-parts
combined_p3_counts <- do.call(rbind, aggregated_counts[p3_subparts]) %>%
  group_by(Target_String) %>%
  summarise(Present_Count = sum(Present_Count)) %>%
  arrange(desc(Present_Count))

# Combine sentiment scores for P3 sub-parts
combined_p3_sentiments <- do.call(rbind, aggregated_sentiments[p3_subparts]) %>%
  group_by(Sentiment) %>%
  summarise(Score = sum(Score)) %>%
  arrange(desc(Score))

# Perform NRC clustering on combined P3 sentiments
combined_p3_nrc_clusters <- Perform_NRC_Clustering(
  data.frame(
    Sentiment = combined_p3_sentiments$Sentiment,
    Score = combined_p3_sentiments$Score,
    stringsAsFactors = FALSE
  )
)


# Print combined results for verification
print(combined_p3_counts)
print(combined_p3_sentiments)
print(combined_p3_nrc_clusters)

# Function to plot aggregated frequency counts as bar graphs
plot_frequency_bar <- function(aggregated_counts, question) {
  ggplot(aggregated_counts, aes(x = reorder(Target_String, -Present_Count), y = Present_Count)) +
    geom_bar(stat = "identity", fill = "skyblue") +
    theme_minimal() +
    labs(title = paste("Frequency of Terms -", question),
         x = "Terms", y = "Count") +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}

# Plot frequency bar graphs for aggregated counts of P1, P2, P3, P4, and P8 (and sub-sections)
plot_frequency_bar(aggregated_counts$P1, "Part 1")
plot_frequency_bar(aggregated_counts$P2_1, "Part 2: Question 1")
plot_frequency_bar(aggregated_counts$P2_2, "Part 2: Question 2")
plot_frequency_bar(combined_p3_counts, "Part 3: Questions 1-5")
plot_frequency_bar(aggregated_counts$P4_1, "Part 4: Question 1")
plot_frequency_bar(aggregated_counts$P4_2, "Part 4: Question 2")
plot_frequency_bar(aggregated_counts$P8_1, "Part 8: Question 1")
plot_frequency_bar(aggregated_counts$P8_2, "Part 8: Question 2")


# Function to plot sentiment analysis results as bar graphs
plot_sentiment_bar <- function(aggregated_sentiment, question) {
  ggplot(aggregated_sentiment, aes(x = reorder(Sentiment, -Score), y = Score, fill = Sentiment)) +
    geom_bar(stat = "identity") +
    theme_minimal() +
    labs(title = paste("Sentiment Analysis -", question),
         x = "Sentiment", y = "Score") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1)) # Rotate x-axis labels for better readability
}

# Plot sentiment bar graphs for aggregated counts of P1, P2, P3, P4, and P8 (and sub-sections)
plot_sentiment_bar(aggregated_sentiments$P1, "Part 1")
plot_sentiment_bar(aggregated_sentiments$P2_1, "Part 2: Question 1")
plot_sentiment_bar(aggregated_sentiments$P2_2, "Part 2: Question 2")
plot_sentiment_bar(combined_p3_sentiments, "Part 3: Questions 1-5")
plot_sentiment_bar(aggregated_sentiments$P4_1, "Part 4: Question 1")
plot_sentiment_bar(aggregated_sentiments$P4_2, "Part 4: Question 2")
plot_sentiment_bar(aggregated_sentiments$P8_1, "Part 8: Question 1")
plot_sentiment_bar(aggregated_sentiments$P8_2, "Part 8: Question 2")

# Function to create a word cloud based on aggregated frequency counts, including combined counts for P3
create_wordcloud <- function(aggregated_counts, part_name, output_dir = "R30_wordclouds") {
  # Create output directory if it doesn't exist
  if (!dir.exists(output_dir)) {
    dir.create(output_dir)
  }
  
  # Define the PNG file name
  png_file <- file.path(output_dir, paste0("wordcloud_", part_name, ".png"))
  
  # Generate the word cloud
  png(png_file, width = 800, height = 800)
  wordcloud(words = aggregated_counts$Target_String, 
            freq = aggregated_counts$Present_Count, 
            scale = c(5, 1),
            colors = brewer.pal(8, "Dark2"))
  dev.off()
}

# Define the parts to visualize
parts_to_plot <- c("P1", "P2_1", "P2_2", "P3_1", "P3_2", "P3_3", "P3_4", "P3_5", "P4_1", "P4_2", "P8_1", "P8_2")

# Generate word clouds for the specified parts
for (part_name in parts_to_plot) {
  create_wordcloud(aggregated_counts[[part_name]], part_name)
}

# Generate a word cloud for combined_p3_counts
create_wordcloud(combined_p3_counts, "combined_P3")

## For generating a word cloud of LDA terms for P1, P3 (combined), P8_1, and P8_2
# Create LDA word cloud for P1
lda_terms_P1 <- unlist(lda_results$P1)
lda_freq_P1 <- table(lda_terms_P1)
lda_df_P1 <- data.frame(Target_String = names(lda_freq_P1), 
                        Present_Count = as.numeric(lda_freq_P1)) %>%
  arrange(desc(Present_Count))

# Create LDA word cloud for P8_1
lda_terms_P8_1 <- unlist(lda_results$P8_1)
lda_freq_P8_1 <- table(lda_terms_P8_1)
lda_df_P8_1 <- data.frame(Target_String = names(lda_freq_P8_1), 
                          Present_Count = as.numeric(lda_freq_P8_1)) %>%
  arrange(desc(Present_Count))

# Create LDA word cloud for P8_2
lda_terms_P8_2 <- unlist(lda_results$P8_2)
lda_freq_P8_2 <- table(lda_terms_P8_2)
lda_df_P8_2 <- data.frame(Target_String = names(lda_freq_P8_2), 
                          Present_Count = as.numeric(lda_freq_P8_2)) %>%
  arrange(desc(Present_Count))


# Combine LDA terms for all sub-parts of P3
combined_lda_terms <- unlist(lapply(c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5"), function(part) {
  unlist(lda_results[[part]])
}))

# Convert the combined LDA terms to a frequency table
lda_terms_freq <- table(combined_lda_terms)

# Filter out terms with frequency less than a threshold
lda_terms_freq <- lda_terms_freq[lda_terms_freq >= 150]

# Convert the frequency table to a data frame
lda_terms_df <- data.frame(Target_String = names(lda_terms_freq), Present_Count = as.numeric(lda_terms_freq))

# Sort the data frame by frequency in descending order
lda_terms_df <- lda_terms_df[order(-lda_terms_df$Present_Count), ]

# Create a word cloud for the combined LDA terms
create_wordcloud <- function(aggregated_counts, part_name, output_dir = "R10_wordclouds") {
  # Create output directory if it doesn't exist
  if (!dir.exists(output_dir)) {
    dir.create(output_dir)
  }
  
  # Define the PNG file name
  png_file <- file.path(output_dir, paste0("wordcloud_", part_name, ".png"))
  
  # Filter out terms with frequency less than a threshold (e.g., 2)
  filtered_counts <- aggregated_counts[aggregated_counts$Present_Count >= 2, ]
  
  # Generate the word cloud (same formatting as others, with a min. frequency argument)
  png(png_file, width = 800, height = 800)
  wordcloud(words = filtered_counts$Target_String, 
            freq = filtered_counts$Present_Count, 
            scale = c(5, 1),
            colors = brewer.pal(8, "Dark2"))
  dev.off()
}

# Generate and save the word cloud for combined LDA terms of P3
create_wordcloud(lda_df_P1, "LDA_P1")
create_wordcloud(lda_df_P8_1, "LDA_P8_1")
create_wordcloud(lda_df_P8_2, "LDA_P8_2")
create_wordcloud(lda_terms_df, "combined_LDA_P3")

# Print the combined LDA terms for verification
print(lda_df_P1)
print(lda_df_P8_1)
print(lda_df_P8_2)
print(lda_terms_df)

######### STATISTICS ##########
####### SPEARMAN'S CORRELATION BETWEEN QUANTITATIVE RATINGS AND FREQUENCY & SENTIMENT RESULTS #######
# Define column mappings for quantitative ratings (Likert-scale of 1-10)
quantitative_columns <- list(
  P1_R1 = 2,     # Part 1 - satisfaction 1: course enjoyment
  P1_R2 = 3,     # Part 1 - satisfaction 2: course meaning (also taps for adaptation with P8_2)
  P1_R3 = 4,     # Part 1 - satisfaction 3: course engagement
  
  P2_R = 10,     # Part 2 - general helpfulness: course helpfulness
  
  P3_1_R1 = 14,  # Part 3 - skill 1 helpfulness
  P3_1_R2 = 15,  # Part 3 - skill 1 usage
  P3_2_R1 = 17,  # Part 3 - skill 2 helpfulness
  P3_2_R2 = 18,  # Part 3 - skill 2 usage
  P3_3_R1 = 20,  # Part 3 - skill 3 helpfulness
  P3_3_R2 = 21,  # Part 3 - skill 3 usage
  P3_4_R1 = 23,  # Part 3 - skill 4 helpfulness
  P3_4_R2 = 24,  # Part 3 - skill 4 usage
  P3_5_R1 = 26,  # Part 3 - skill 5 helpfulness
  P3_5_R2 = 27   # Part 3 - skill 5 usage
)

# Extract the quantitative ratings for each participant
quantitative_ratings <- df %>%
  select(anonymous_id, all_of(unlist(quantitative_columns)))

# Rename columns for organization
colnames(quantitative_ratings) <- c("anonymous_id", names(quantitative_columns))

# Print the extracted ratings for verification
print(quantitative_ratings)

## Aggregate frequency count by participant for each part ##
# Create a list to store aggregated counts by participant
aggregated_counts_by_participant <- list()

for (part in names(count_tables_list)) {
  # Combine all participant data for the current part
  part_data <- bind_rows(count_tables_list[[part]])
  
  # Aggregate Present_Count by anonymous_id
  aggregated_counts_by_participant[[part]] <- part_data %>%
    group_by(anonymous_id) %>%
    summarise(Present_Count = sum(Present_Count, na.rm = TRUE))
}

## Merge quantitative_ratings with aggregated_counts_by_participant for each part ##
# Create a list to store merged data for each part
merged_frequency_data <- list()

# Merge for P1
merged_frequency_data$P1 <- quantitative_ratings %>%
  select(anonymous_id, P1_R1, P1_R2, P1_R3) %>%
  left_join(aggregated_counts_by_participant[["P1"]], by = "anonymous_id")

# Merge for P2_1 and P2_2
merged_frequency_data$P2_1 <- quantitative_ratings %>%
  select(anonymous_id, P2_R) %>%
  left_join(aggregated_counts_by_participant[["P2_1"]], by = "anonymous_id")

merged_frequency_data$P2_2 <- quantitative_ratings %>%
  select(anonymous_id, P2_R) %>%
  left_join(aggregated_counts_by_participant[["P2_2"]], by = "anonymous_id")

# Merge for P3_1 to P3_5
for (part in c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")) {
  merged_frequency_data[[part]] <- quantitative_ratings %>%
    select(anonymous_id, paste0(part, "_R1"), paste0(part, "_R2")) %>%
    left_join(aggregated_counts_by_participant[[part]], by = "anonymous_id")
}

# Merge for P8_2
merged_frequency_data$P8_2 <- quantitative_ratings %>%
  select(anonymous_id, P1_R2) %>%
  left_join(aggregated_counts_by_participant[["P8_2"]], by = "anonymous_id")

#### RUNNING SPEARMAN'S R ON RATINGS & FREQUENCY SCORES ####
# Initialize lists to store results (correlation & descriptive statistics)
frequency_cor_results <- list()
desc_stats_list <- list()

# Function to calculate robust descriptive statistics
get_desc_stats <- function(x, name) {
  stats <- list(
    Variable = name,
    N = sum(!is.na(x)),
    Median = round(median(x, na.rm = TRUE), 2),
    IQR = round(IQR(x, na.rm = TRUE), 2),
    Min = round(min(x, na.rm = TRUE), 2),
    Max = round(max(x, na.rm = TRUE), 2)
  )
  return(stats)
}

## P1 CORRELATIONS ##
# Descriptive stats for P1 variables
desc_stats_list$P1_R1 <- get_desc_stats(merged_frequency_data$P1$P1_R1, "P1_R1 (Enjoyment)")
desc_stats_list$P1_R2 <- get_desc_stats(merged_frequency_data$P1$P1_R2, "P1_R2 (Meaning)")
desc_stats_list$P1_R3 <- get_desc_stats(merged_frequency_data$P1$P1_R3, "P1_R3 (Engagement)")
desc_stats_list$P1_Freq <- get_desc_stats(merged_frequency_data$P1$Present_Count, "P1 Frequency")

# correlation tests
frequency_cor_results$P1_R1 <- cor.test(merged_frequency_data$P1$P1_R1, merged_frequency_data$P1$Present_Count, method = "spearman", exact = FALSE)
frequency_cor_results$P1_R2 <- cor.test(merged_frequency_data$P1$P1_R2, merged_frequency_data$P1$Present_Count, method = "spearman", exact = FALSE)
frequency_cor_results$P1_R3 <- cor.test(merged_frequency_data$P1$P1_R3, merged_frequency_data$P1$Present_Count, method = "spearman", exact = FALSE)

## P2 CORRELATIONS ##
# Descriptive stats for P2 variables
desc_stats_list$P2_R <- get_desc_stats(merged_frequency_data$P2_1$P2_R, "P2_R (Helpfulness)")
desc_stats_list$P2_1_Freq <- get_desc_stats(merged_frequency_data$P2_1$Present_Count, "P2_1 Frequency")
desc_stats_list$P2_2_Freq <- get_desc_stats(merged_frequency_data$P2_2$Present_Count, "P2_2 Frequency")

# correlation tests
frequency_cor_results$P2_R_P2_1 <- cor.test(merged_frequency_data$P2_1$P2_R, merged_frequency_data$P2_1$Present_Count, method = "spearman", exact = FALSE)
frequency_cor_results$P2_R_P2_2 <- cor.test(merged_frequency_data$P2_2$P2_R, merged_frequency_data$P2_2$Present_Count, method = "spearman", exact = FALSE)

## P3 CORRELATIONS ##
for (part in c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")) {
  # Get data
  rating_R1 <- merged_frequency_data[[part]][[paste0(part, "_R1")]]
  rating_R2 <- merged_frequency_data[[part]][[paste0(part, "_R2")]]
  present_count <- merged_frequency_data[[part]]$Present_Count
  
  # Descriptive stats
  desc_stats_list[[paste0(part, "_R1")]] <- get_desc_stats(rating_R1, paste0(part, "_R1 (Helpfulness)"))
  desc_stats_list[[paste0(part, "_R2")]] <- get_desc_stats(rating_R2, paste0(part, "_R2 (Usage)"))
  desc_stats_list[[paste0(part, "_Freq")]] <- get_desc_stats(present_count, paste0(part, " Frequency"))
  
  # correlation tests
  frequency_cor_results[[paste0(part, "_R1")]] <- cor.test(
    rating_R1, 
    present_count, 
    method = "spearman", 
    exact = FALSE
  )
  frequency_cor_results[[paste0(part, "_R2")]] <- cor.test(
    rating_R2, 
    present_count, 
    method = "spearman", 
    exact = FALSE
  )
}

## P8_2 CORRELATION ##
# Descriptive stats
desc_stats_list$P8_2_R2 <- get_desc_stats(merged_frequency_data$P8_2$P1_R2, "P8_2_R2 (Meaning)")
desc_stats_list$P8_2_Freq <- get_desc_stats(merged_frequency_data$P8_2$Present_Count, "P8_2 Frequency")

# correlation test
frequency_cor_results$P8_2 <- cor.test(merged_frequency_data$P8_2$P1_R2, merged_frequency_data$P8_2$Present_Count, method = "spearman", exact = FALSE)

## RESULTS OUTPUT ##
# Convert descriptive stats to clean data frame
desc_stats_df <- do.call(rbind, lapply(desc_stats_list, as.data.frame))
rownames(desc_stats_df) <- NULL

# correlation results processing
frequency_cor_results_df <- data.frame(
  Correlation = names(frequency_cor_results),
  Spearman_R = round(sapply(frequency_cor_results, function(x) x$estimate), 3),
  p_value = format.pval(sapply(frequency_cor_results, function(x) x$p.value), digits = 3)
)

# Print results with clear separation
cat("\n\nDESCRIPTIVE STATISTICS\n")
print(desc_stats_df)

cat("\n\nSPEARMAN'S RANK CORRELATIONS\n")
print(frequency_cor_results_df)

#### RUNNING SPEARMAN'S R ON RATINGS & SENTIMENT SCORES ####
# Create a list to store aggregated sentiment scores by participant
aggregated_sentiment_by_participant <- list()

for (part in names(sentiment_results)) {
  # Aggregate sentiment scores by anonymous_id
  aggregated_sentiment_by_participant[[part]] <- sentiment_results[[part]] %>%
    group_by(anonymous_id) %>%
    summarise(across(everything(), sum, na.rm = TRUE))
}

# Create a list to store merged data for each part
merged_sentiment_data <- list()

# Merge for P1
merged_sentiment_data$P1 <- quantitative_ratings %>%
  select(anonymous_id, P1_R1, P1_R2, P1_R3) %>%
  left_join(aggregated_sentiment_by_participant[["P1"]], by = "anonymous_id")

# Merge for P2_1 and P2_2
merged_sentiment_data$P2_1 <- quantitative_ratings %>%
  select(anonymous_id, P2_R) %>%
  left_join(aggregated_sentiment_by_participant[["P2_1"]], by = "anonymous_id")

merged_sentiment_data$P2_2 <- quantitative_ratings %>%
  select(anonymous_id, P2_R) %>%
  left_join(aggregated_sentiment_by_participant[["P2_2"]], by = "anonymous_id")

# Merge for P3_1 to P3_5
for (part in c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")) {
  merged_sentiment_data[[part]] <- quantitative_ratings %>%
    select(anonymous_id, paste0(part, "_R1"), paste0(part, "_R2")) %>%
    left_join(aggregated_sentiment_by_participant[[part]], by = "anonymous_id")
}

# Merge for P8_2
merged_sentiment_data$P8_2 <- quantitative_ratings %>%
  select(anonymous_id, P1_R2) %>%
  left_join(aggregated_sentiment_by_participant[["P8_2"]], by = "anonymous_id")

#### RUNNING SPEARMAN'S R ON RATINGS & SENTIMENT SCORES ####
# Initialize lists to store results
sentiment_cor_results <- list()
sentiment_desc_stats_list <- list()  # Unique name for sentiment stats

# Reuse the same descriptive stats function from previous frequency section
get_desc_stats <- function(x, name) {
  stats <- list(
    Variable = name,
    N = sum(!is.na(x)),
    Median = round(median(x, na.rm = TRUE), 2),
    IQR = round(IQR(x, na.rm = TRUE), 2),
    Min = round(min(x, na.rm = TRUE), 2),
    Max = round(max(x, na.rm = TRUE), 2)
  )
  return(stats)
}

## P1 SENTIMENT CORRELATIONS ##
# Descriptive stats for P1 sentiment variables
sentiment_desc_stats_list$P1_R1 <- get_desc_stats(merged_sentiment_data$P1$P1_R1, "P1_R1 (Enjoyment)")
sentiment_desc_stats_list$P1_R2 <- get_desc_stats(merged_sentiment_data$P1$P1_R2, "P1_R2 (Meaning)")
sentiment_desc_stats_list$P1_R3 <- get_desc_stats(merged_sentiment_data$P1$P1_R3, "P1_R3 (Engagement)")
sentiment_desc_stats_list$P1_Sentiment <- get_desc_stats(rowSums(merged_sentiment_data$P1[, -1]), "P1 Sentiment")

# Correlation tests (unchanged)
sentiment_cor_results$P1_R1 <- cor.test(merged_sentiment_data$P1$P1_R1, rowSums(merged_sentiment_data$P1[, -1]), method = "spearman", exact = FALSE)
sentiment_cor_results$P1_R2 <- cor.test(merged_sentiment_data$P1$P1_R2, rowSums(merged_sentiment_data$P1[, -1]), method = "spearman", exact = FALSE)
sentiment_cor_results$P1_R3 <- cor.test(merged_sentiment_data$P1$P1_R3, rowSums(merged_sentiment_data$P1[, -1]), method = "spearman", exact = FALSE)

## P2 SENTIMENT CORRELATIONS ##
# Descriptive stats for P2 sentiment variables
sentiment_desc_stats_list$P2_R <- get_desc_stats(merged_sentiment_data$P2_1$P2_R, "P2_R (Helpfulness)")
sentiment_desc_stats_list$P2_1_Sentiment <- get_desc_stats(rowSums(merged_sentiment_data$P2_1[, -1]), "P2_1 Sentiment")
sentiment_desc_stats_list$P2_2_Sentiment <- get_desc_stats(rowSums(merged_sentiment_data$P2_2[, -1]), "P2_2 Sentiment")

# Correlation tests
sentiment_cor_results$P2_R_P2_1 <- cor.test(merged_sentiment_data$P2_1$P2_R, rowSums(merged_sentiment_data$P2_1[, -1]), method = "spearman", exact = FALSE)
sentiment_cor_results$P2_R_P2_2 <- cor.test(merged_sentiment_data$P2_2$P2_R, rowSums(merged_sentiment_data$P2_2[, -1]), method = "spearman", exact = FALSE)

## P3 SENTIMENT CORRELATIONS ##
for (part in c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")) {
  # Get data
  rating_R1 <- merged_sentiment_data[[part]][[paste0(part, "_R1")]]
  rating_R2 <- merged_sentiment_data[[part]][[paste0(part, "_R2")]]
  sentiment_score <- rowSums(merged_sentiment_data[[part]][, -1])
  
  # Descriptive stats
  sentiment_desc_stats_list[[paste0(part, "_R1")]] <- get_desc_stats(rating_R1, paste0(part, "_R1 (Helpfulness)"))
  sentiment_desc_stats_list[[paste0(part, "_R2")]] <- get_desc_stats(rating_R2, paste0(part, "_R2 (Usage)"))
  sentiment_desc_stats_list[[paste0(part, "_Sentiment")]] <- get_desc_stats(sentiment_score, paste0(part, " Sentiment"))
  
  # Correlation tests
  sentiment_cor_results[[paste0(part, "_R1")]] <- cor.test(
    rating_R1, 
    sentiment_score, 
    method = "spearman", 
    exact = FALSE
  )
  sentiment_cor_results[[paste0(part, "_R2")]] <- cor.test(
    rating_R2, 
    sentiment_score, 
    method = "spearman", 
    exact = FALSE
  )
}

## P8_2 SENTIMENT CORRELATION ##
# Descriptive stats
sentiment_desc_stats_list$P8_2_R2 <- get_desc_stats(merged_sentiment_data$P8_2$P1_R2, "P8_2_R2 (Meaning)")
sentiment_desc_stats_list$P8_2_Sentiment <- get_desc_stats(rowSums(merged_sentiment_data$P8_2[, -1]), "P8_2 Sentiment")

# Correlation test
sentiment_cor_results$P8_2 <- cor.test(merged_sentiment_data$P8_2$P1_R2, rowSums(merged_sentiment_data$P8_2[, -1]), method = "spearman", exact = FALSE)

## SENTIMENT RESULTS OUTPUT ##
# Convert descriptive stats to clean data frame (using unique name)
sentiment_desc_stats_df <- do.call(rbind, lapply(sentiment_desc_stats_list, as.data.frame))
rownames(sentiment_desc_stats_df) <- NULL

# Correlation results processing
sentiment_cor_results_df <- data.frame(
  Correlation = names(sentiment_cor_results),
  Spearman_R = round(sapply(sentiment_cor_results, function(x) x$estimate), 3),
  p_value = format.pval(sapply(sentiment_cor_results, function(x) x$p.value), digits = 3)
)

# Print results with clear labels
cat("\n\nSENTIMENT DESCRIPTIVE STATISTICS\n")
print(sentiment_desc_stats_df)

cat("\n\nSENTIMENT SPEARMAN'S RANK CORRELATIONS\n")
print(sentiment_cor_results_df)

#### RUNNING SPEARMAN'S R BETWEEN FREQUENCY & SENTIMENT SCORES ####
# Initialize list to store results
freq_sent_cor_results <- list()

# P1 correlation
freq_sent_cor_results$P1 <- cor.test(
  merged_frequency_data$P1$Present_Count,
  rowSums(merged_sentiment_data$P1[, -1]),
  method = "spearman",
  exact = FALSE
)

# P2 correlations
freq_sent_cor_results$P2_1 <- cor.test(
  merged_frequency_data$P2_1$Present_Count,
  rowSums(merged_sentiment_data$P2_1[, -1]),
  method = "spearman",
  exact = FALSE
)

freq_sent_cor_results$P2_2 <- cor.test(
  merged_frequency_data$P2_2$Present_Count,
  rowSums(merged_sentiment_data$P2_2[, -1]),
  method = "spearman",
  exact = FALSE
)

# P3 correlations
for (part in c("P3_1", "P3_2", "P3_3", "P3_4", "P3_5")) {
  freq_sent_cor_results[[part]] <- cor.test(
    merged_frequency_data[[part]]$Present_Count,
    rowSums(merged_sentiment_data[[part]][, -1]),
    method = "spearman",
    exact = FALSE
  )
}

# P8_2 correlation
freq_sent_cor_results$P8_2 <- cor.test(
  merged_frequency_data$P8_2$Present_Count,
  rowSums(merged_sentiment_data$P8_2[, -1]),
  method = "spearman",
  exact = FALSE
)

# Correlation results processing
freq_sent_cor_results_df <- data.frame(
  Correlation = names(freq_sent_cor_results),
  Spearman_R = round(sapply(freq_sent_cor_results, function(x) x$estimate), 3),
  p_value = format.pval(sapply(freq_sent_cor_results, function(x) x$p.value), digits = 3)
)

# Print results with clear labels
cat("\n\nFREQUENCY & SENTIMENT SPEARMAN'S RANK CORRELATIONS\n")
print(freq_sent_cor_results_df)

###################### EXPORTING RESULTS TO EXCEL ######################
#Exporting individual results for questions all Parts (and sub-parts) to Excel
output_file_path <- "July_30_Thesis_Individual_Results_Tables.xlsx"

# Create a new workbook
wb <- createWorkbook()

# Add sheets for frequency, sentiment, NRC clustering, and LDA results for each part
parts <- c("P1", "P2_1", "P2_2", "P3_1", "P3_2", "P3_3", "P3_4", "P3_5", "P4_1", "P4_2", "P8_1", "P8_2")

# Loop through each part to add worksheets and write data
for (part in parts) {
  # Add worksheets for each part
  addWorksheet(wb, paste0(part, "_Frequency_Results"))
  addWorksheet(wb, paste0(part, "_Sentiment_Results"))
  addWorksheet(wb, paste0(part, "_NRC_Cluster_Results"))
  addWorksheet(wb, paste0(part, "_LDA_Results"))
  
  # Write frequency results
  writeData(wb, paste0(part, "_Frequency_Results"), count_tables_list[[part]])
  
  # Write sentiment results
  writeData(wb, paste0(part, "_Sentiment_Results"), sentiment_results[[part]])
  
  # Write NRC clustering results
  writeData(wb, paste0(part, "_NRC_Cluster_Results"), nrc_cluster_results[[part]])
  
  # Write LDA results
  lda_data_list <- lapply(names(lda_results[[part]]), function(participant_id) {
    topic_terms <- lda_results[[part]][[participant_id]]
    if (!is.null(topic_terms)) {
      # Convert topic terms to a data frame
      df_topic_terms <- as.data.frame(topic_terms, stringsAsFactors = FALSE)
      # Add participant ID as a column
      df_topic_terms$anonymous_id <- participant_id
      return(df_topic_terms)
    }
  })
  
  # Find all unique column names across participants' LDA results
  all_columns <- unique(unlist(lapply(lda_data_list, colnames)))
  
  # Standardize column names and order for all participants
  lda_data_list <- lapply(lda_data_list, function(df) {
    # Add missing columns with NA values
    missing_columns <- setdiff(all_columns, colnames(df))
    for (col in missing_columns) {
      df[[col]] <- NA
    }
    # Reorder columns to match the standard order
    df <- df[, all_columns, drop = FALSE]
    return(df)
  })
  
  # Combine all participants' LDA results into a single data frame
  lda_data <- do.call(rbind, lda_data_list)
  
  # Write LDA results to the worksheet
  if (!is.null(lda_data)) {
    writeData(wb, paste0(part, "_LDA_Results"), lda_data)
  }
}

# Add worksheets for sentiment shift analyses
addWorksheet(wb, "Sentiment_Shift_P2")
addWorksheet(wb, "Sentiment_Shift_P8")

# Write sentiment shift analysis results
writeData(wb, "Sentiment_Shift_P2", sentiment_shift_P2)
writeData(wb, "Sentiment_Shift_P8", sentiment_shift_P8)

#add worksheets for combined P3 NLP analyses
addWorksheet(wb, "Combined_P3_Frequency")
addWorksheet(wb, "Combined_P3_Sentiment")
addWorksheet(wb, "Combined_P3_NRC_Clusters")

# Write combined P3 results
writeData(wb, "Combined_P3_Frequency", combined_p3_counts)
writeData(wb, "Combined_P3_Sentiment", combined_p3_sentiments)
writeData(wb, "Combined_P3_NRC_Clusters", combined_p3_nrc_clusters)

# Add a worksheet for quantitative ratings
addWorksheet(wb, "Quantitative_Ratings")

# Write quantitative ratings to the worksheet
writeData(wb, "Quantitative_Ratings", quantitative_ratings)

# Add worksheets for Sentiment & Frequency descriptive statistics
addWorksheet(wb, "Freq_Descriptive_Stats")
addWorksheet(wb, "Sent_Descriptive_Stats")

# Write descriptive statistics
writeData(wb, "Freq_Descriptive_Stats", desc_stats_df)
writeData(wb, "Sent_Descriptive_Stats", sentiment_desc_stats_df)

# Add worksheets for Spearman's correlation results (frequency and sentiment)
addWorksheet(wb, "Frequency_Correlation_Results")
addWorksheet(wb, "Sentiment_Correlation_Results")
addWorksheet(wb, "Frequency_Sentiment_Cor_Results")

# Write correlation results
writeData(wb, "Frequency_Correlation_Results", frequency_cor_results_df)
writeData(wb, "Sentiment_Correlation_Results", sentiment_cor_results_df)
writeData(wb, "Frequency_Sentiment_Cor_Results", freq_sent_cor_results_df)

# Save the workbook
saveWorkbook(wb, output_file_path, overwrite = TRUE)
cat("Excel file saved at:", output_file_path)

###################### END OF SCRIPT ######################

dev.off()