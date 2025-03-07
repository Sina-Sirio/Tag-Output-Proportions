import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
df = pd.read_excel("my_output_with_topics.xlsx")

# Ensure "Topic" column exists
if "Topic" not in df.columns:
    raise ValueError("The 'Topic' column is missing from the Excel file.")

# Count occurrences of each unique value in "Topic"
topic_counts = df["Topic"].value_counts()

topic_counts_df = pd.DataFrame({"Topic": topic_counts.index, "Count": topic_counts.values})

# Save the dataframe to a CSV file
topic_counts_df.to_csv("topic_counts.csv", index=False)
print("Topic occurrences saved to topic_counts.csv")

# Compute proportions
topic_proportions = topic_counts / topic_counts.sum()
topic_proportions_df = pd.DataFrame({"Topic": topic_counts.index, "Proportion": topic_proportions.values})

# Save proportions to CSV
topic_proportions_df.to_csv("topic_proportions.csv", index=False)
print("Topic proportions saved to topic_proportions.csv")

# Set style for seaborn
sns.set_style("whitegrid")

# Plot bar chart
plt.figure(figsize=(12, 6))
sns.barplot(x=topic_counts.index, y=topic_counts.values, palette="coolwarm")
plt.xlabel("Topic")
plt.ylabel("Count")
plt.title("Occurrences of Each Unique Topic")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Plot pie chart
plt.figure(figsize=(8, 8))
topic_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, cmap='tab10')
plt.ylabel("")
plt.title("Proportion of Each Unique Topic")
plt.show()

# Plot a donut chart
plt.figure(figsize=(8, 8))
wedges, texts, autotexts = plt.pie(topic_counts, autopct='%1.1f%%', startangle=90, colors=sns.color_palette("pastel"), wedgeprops=dict(width=0.4))
plt.setp(autotexts, size=10, weight="bold")
plt.title("Donut Chart - Topic Distribution")
plt.show()

# Plot a cumulative percentage graph
plt.figure(figsize=(12, 6))
cumulative_counts = topic_counts.cumsum() / topic_counts.sum() * 100
sns.lineplot(x=topic_counts.index, y=cumulative_counts, marker="o", linestyle="--", color="b")
plt.xlabel("Topic")
plt.ylabel("Cumulative Percentage (%)")
plt.title("Cumulative Distribution of Topics")
plt.xticks(rotation=45, ha='right')
plt.grid(True)
plt.show()