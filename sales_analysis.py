# Import  libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset 
df = pd.read_excel("Sales Report.xlsx")


print(df.head())


print(df.isnull().sum())

# Fill missing values or drop
df = df.dropna()  

#describe statistics
print(df.describe())

# Top 10 products by sales
top_products = df.groupby('Product')['Sales'].sum().sort_values(ascending=False).head(10)
print(top_products)

# Plot  products
plt.figure(figsize=(10,6))
sns.barplot(x=top_products.values, y=top_products.index, palette="viridis")
plt.title("Top 10 Products by Sales")
plt.xlabel("Total Sales")
plt.ylabel("Product")
plt.show()

# Monthly sales 
df['Order_Date'] = pd.to_datetime(df['Order_Date'])
df['Month'] = df['Order_Date'].dt.to_period('M')
monthly_sales = df.groupby('Month')['Sales'].sum()

plt.figure(figsize=(12,6))
monthly_sales.plot(kind='line', marker='o')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.show()

category_sales = df.groupby('Category')['Sales'].sum()

plt.figure(figsize=(8,6))
sns.barplot(x=category_sales.index, y=category_sales.values, palette="pastel")
plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Total Sales")
plt.show()
