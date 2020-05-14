# -*- coding: utf-8 -*-
import pandas as pd
xl=pd.ExcelFile("SalesData.xlsx")
SalesData = xl.parse("Orders")
import matplotlib.pyplot as plt
import seaborn as sns
import sqlite3
conn = sqlite3.connect('OS_Employee.db')
xl =  pd.ExcelFile("SalesData.xlsx")
Decorating = "\n" + "*"*45 + "\n"
cur = conn.cursor()
 
def MostProfitableSubCategory():
    import pandas as pd
    xl = pd.ExcelFile("SalesData.xlsx")
    SalesData = xl.parse("Orders")
    Decorating = "\n" + "*"*25 + "\n"
    
    categories=SalesData.Category.unique()
    
    #Top 5 Most Profitable Sub-Category by Category      
    
    for category in categories:
        cat_select=SalesData.loc[SalesData["Category"]==category]
        cat_subcat_prof=cat_select[["Category","Sub-Category","Profit"]]
        most_sub = cat_subcat_prof.groupby("Sub-Category").sum().sort_values(by="Profit", ascending = False)
        print("\nMost Profitable Sub-Category in the {} category: ".format(category))
        print(most_sub.head(5))
        print(Decorating)
        
def LeastProfitableSubCategory():
        # -*- coding: utf-8 -*-
    # -*- coding: utf-8 -*-
    import pandas as pd
    xl = pd.ExcelFile("SalesData.xlsx")
    SalesData = xl.parse("Orders")
    Decorating = "\n" + "*"*25 + "\n"
    
    categories=SalesData.Category.unique()
    
    #Top 5 Least Profitable Sub-Category by Category
           
          
    
    for category in categories:
        cat_select=SalesData.loc[SalesData["Category"]==category]
        cat_subcat_prof=cat_select[["Category","Sub-Category","Profit"]]
        least_sub = cat_subcat_prof.groupby("Sub-Category").sum().sort_values(by="Profit", ascending = True)
        print("\nLeast Profitable Sub-Category in the {} category: ".format(category))
        print(least_sub.head(5))
        print(Decorating)
        
def MostProfitableSubCategoryByRegion():
        # -*- coding: utf-8 -*-
    # -*- coding: utf-8 -*-
    import pandas as pd
    xl = pd.ExcelFile("SalesData.xlsx")
    SalesData = xl.parse("Orders")
    Decorating = "\n" + "*"*25 + "\n"
    
    categories=SalesData.Category.unique()
    regions=SalesData.Region.unique()
    
    #Top 5 Most Profitable Sub-Category by Category in each region
           
          
    
    for category in categories:
        cat_select=SalesData.loc[SalesData["Category"]==category]
        cat_sub_prof_reg=cat_select[["Category","Sub-Category","Profit","Region"]]
        for region in regions:
            best_sub_reg = cat_sub_prof_reg.loc[cat_sub_prof_reg["Region"]==region]
            best_sub = best_sub_reg.groupby("Sub-Category").sum().sort_values(by="Profit", ascending = False)
            print("\nMost Profitable Sub-Category in the {} region for {} category: ".format(region,category))
            print(best_sub.head(10))
            print(Decorating)
            
def LeastProfitableSubCategoryByRegion():
        # -*- coding: utf-8 -*-
    # -*- coding: utf-8 -*-
    import pandas as pd
    xl = pd.ExcelFile("SalesData.xlsx")
    SalesData = xl.parse("Orders")
    Decorating = "\n" + "*"*25 + "\n"
    
    categories=SalesData.Category.unique()
    regions=SalesData.Region.unique()
    
    #Least Profitable Sub-Category by Category in each region
           
          
    
    for category in categories:
        cat_select=SalesData.loc[SalesData["Category"]==category]
        cat_sub_prof_reg=cat_select[["Category","Sub-Category","Profit","Region"]]
        for region in regions:
            best_sub_reg = cat_sub_prof_reg.loc[cat_sub_prof_reg["Region"]==region]
            best_sub = best_sub_reg.groupby("Sub-Category").sum().sort_values(by="Profit", ascending = True)
            print("\nLeast Profitable Sub-Category in the {} region for {} category: ".format(region,category))
            print(best_sub.head(10))
            print(Decorating)

def SubCategoryAnalytics():
    print(Decorating)
    print("You are in Sub-Category Analytics Sub-Menu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Most Profitable Sub-Category by Category \n 2: Least Profitable Sub-Category by Category")
    print(" 3: Most Profitable Sub-Category by Region \n 4: Least Profitable Sub-Category by Region \n 5: Return to Main Menu" )
    
    print(Decorating)
    Choice = input("Please select an option: ")
    #Defines and displays the most profitable sub-category by category
    
    if Choice == "1":
        MostProfitableSubCategory()
        input("Press enter to go back to Sub-Category Analytics:")
       #Returns to the Segment Analytics Menu
        SubCategoryAnalytics()            
    #Displays the least profitable sub-category by category
    elif Choice == "2":
        LeastProfitableSubCategory()
        input("Press enter to go back to Sub-Category Analytics:")
       #Returns to the Segment Analytics Menu
        SubCategoryAnalytics() 
        
        #Top 5 Least Profitable Sub-Category by Category

        
        #Displays most profitable sub-category by region
    elif Choice == "3":
        MostProfitableSubCategoryByRegion()
        input("Press enter to go back to Sub-Category Analytics:")
       #Returns to the Segment Analytics Menu
        SubCategoryAnalytics() 
        
    elif Choice == "4":
        LeastProfitableSubCategoryByRegion()
        input("Press enter to go back to Sub-Category Analytics:")
       #Returns to the Segment Analytics Menu
        SubCategoryAnalytics() 
#Provides analytic visuals for Discount Analytics
def DiscountAnalytics():
    print(Decorating)
    print("You are in Discount Analytics Sub-Menu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Total Profits by Discount Rates \n 2: Maximum and Minimum Discount Per Sub-Category")
    print(" 3: Sales by Discount Correlation \n 4: Profits by Discount Correlation \n 5: Return to Main Menu" )
    
    print(Decorating)
    Choice = input("Please select an option: ")
    #Defines and displays the total profit by discount
    
    if Choice == "1":
        print("Total Profits by Discount:")
        Profits_by_Discount = SalesData.groupby(["Discount"]).agg({'Profit':sum})
        print(Profits_by_Discount)   
        input("Press enter to go back to Discount Analytics:")
       #Returns to Main Menu
        DiscountAnalytics()
        
    #Displays the maximum and minimum discount rate for each sub-category
    elif Choice == "2":
   #gives the maximum rate for each sub-category
       Subcategories_BY_Discount = SalesData[["Sub-Category","Discount"]]
       Subcategory_Max_Rate = Subcategories_BY_Discount.groupby(by="Sub-Category").max()
       print("Maximum and Minimum Discount Rate for each Sub-Category:")
       print(Decorating)
       print(Subcategory_Max_Rate)
       print(Decorating)
       #Provides space between this and the next insight component
       print()
       #gives the minimum rate for each sub-category
       Subcategory_Min_Rate = Subcategories_BY_Discount.groupby(by="Sub-Category").min()
       print(Decorating)
       print(Subcategory_Min_Rate)
       print(Decorating)
    
       input("Press enter to go back to Discount Analytics:")
       #Returns to DiscountAnalytics
       DiscountAnalytics()
    #Displays the visuals that show the correlation between sales and discounts
    elif Choice == "3":
        #Shows the line chart between discount and sales
        sns.regplot(x = "Discount", y = "Sales", fit_reg = True, data = SalesData)
        plt.xlabel("Discount Rate")
        plt.ylabel("Sales")
        plt.title("Relationship between Sales and Discount \n")
        plt.show()
            
        input("Press enter to go back to Discount Analytics:")
       #Returns to Discount Analytics Menu
        DiscountAnalytics()
    
#    
#Displays the visuals that show the correlation between profits and discounts
    elif Choice == "4":
        #Helps create the line chart that shows correlation between profits and discount rates
        sns.regplot(x = "Discount", y = "Profit", fit_reg = True, data = SalesData)
        plt.xlabel("Discount Rate")
        plt.ylabel("Profit")
        plt.title("Relationship between Profit and Discount \n")
        plt.show()
        
        input("Press enter to go back to Discount Analytics:")
       #Returns to Discount Analytics Menu
        DiscountAnalytics()

    

    elif Choice == "5":
        #Helps create the line chart that shows correlation between profits and discount rates
        #Returns to Main Menu
        
        MainMenu()
        
       #Provides space between the two insight components
        
    else:
        print("Error! Please select an option between numbers:1-5")
        input("Press enter to go back to Discount Analytics:")
       #Returns to Main Menu
        DiscountAnalytics()

def StateAnalytics():
    print(Decorating)
    print("You are in State Analytics Sub-Menu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Top 10 States by Sales \n 2: Worst 10 States by Sales")
    print(" 3: Top 10 States by Profit \n 4: Worst 10 States by Profit \n 5: Average Discount Rate by State")
    print(" 6: Return to the Main Menu")
     
    print(Decorating)
    Choice = input("Please select on option: ")
    
    State_sales = SalesData [["State", "Sales"]]
    State_total_Sales_Profit = State_sales.groupby(by="State").sum().sort_values(by="Sales", ascending = False)
    State_profit = SalesData [["State", "Profit"]]
    State_total_Profit = State_profit.groupby(by="State").sum().sort_values(by="Profit", ascending = False)
    State_Discount = SalesData [["State", "Discount"]]
    Avg_Discount_Per_State = (State_Discount.groupby(["State"],as_index=False).mean().groupby("State")["Discount"].mean())

    #Displays the top 15 states by total sales in the descending order     
    if Choice == "1":
        print(Decorating)  
        print("Top 15 States by Sales:")
        print(State_total_Sales_Profit.head(15))
        print(Decorating)
        input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
        StateAnalytics()
        

    #Displays the worst 15 states by total sales in the ascending order
    elif Choice == "2":
        print(Decorating)
        print("Worst 15 States by Sales:")        
        print(State_total_Sales_Profit.tail(15))  
        input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
        StateAnalytics()
    
    
    #Displays the worst 15 states by total profit in the descending order
    elif Choice == "4":
        print(Decorating)
        print("Top 15 States by Sales:")        
        print(State_total_Profit.tail(15))  
        print(Decorating)
        input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
        StateAnalytics()      
     
    
    
    
   #Displays the top 15 states by total profit in ascending order
    elif Choice == "3":
        print(Decorating)
        print("Worst 15 States by Sales:")                
        print(State_total_Profit.head(15))  
        print(Decorating)
        input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
        StateAnalytics()
   
    elif Choice == "5":
        print(Decorating)
        print("States and their average discount rates:")        
        print(Avg_Discount_Per_State)         
        print(Decorating)
        input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
        StateAnalytics()
        
    elif Choice == "6":
#       #Helps create the line chart that shows correlation between profits and discount rates
        #Returns to Main Menu
        MainMenu()
#        
#        #Provides space between the two insight components
#        print()

    else:
       print("Error! Please select an option between numbers:1-6")
       input("Press enter to go back to State Analytics:")
       #Returns to Main Menu
       StateAnalytics()

def SegmentAnalytics():
    #Displays the customer segments and their total profits
    Segment_profit = SalesData [["Segment", "Profit"]]
    Segment_total_Profit = Segment_profit.groupby(by="Segment").sum().sort_values(by="Profit", ascending = False)
    #Defines the properties for the chart
    cust_seg_info = SalesData[["Segment","Sales"]]
    cust_seg_sales = SalesData.groupby(["Segment"]).sum().sort_values(by=["Sales"]).reset_index()
    designations = list(cust_seg_sales.Segment.values)
    sales_total = list(cust_seg_sales.Sales.values)
    #Shows the Total Profit by Product Name and Segment
    Product_Segment_Profit_info = SalesData[["Product Name","Segment", "Profit"]]
    #Shows the best products by profit and segment
    Total_Profit_By_Segment_Product = Product_Segment_Profit_info.groupby(["Product Name", "Segment"]).sum().sort_values(by=['Profit'])
    
    #Tells the user sub-menu options
    print(Decorating)
    print("You are in Segment Analytics Sub-Menu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Ranking each segment by total profit \n 2: Total Sales by each segment")
    print(" 3: The top 10 products by segment and total profit \n 4: The worst 10 products by segment and total profit \n 5: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select on option: ")
    #Displays the top 3 records 
    if Choice == "1":
        print("Segments by Total Profit")
        print(Segment_total_Profit.head(3))
        input("Press Enter to go back to Segment Analytics:")
       #Returns to the Segment Analytics Menu
        SegmentAnalytics()



    #Prints the Pie Chart by segment and sales percentage
    elif Choice == "2":
        print(Decorating)
        print("Segments by Total Sales Percentage:")
        print()
        print(designations)
        print(sales_total)
        #Gives the colors of pie chart
        symbols = ['cyan','red','orange']
        explode = (0, 0, 0.1)
        sales_pie_chart = plt.pie(sales_total, labels = designations, explode=explode,shadow= True, colors=symbols,autopct='%1.1f%%')
        plt.axis('equal')
        plt.title("Total Sales Percentage per Customer Segment\n")
        plt.show()
        print(Decorating)
        input("Press Enter to go back to Segment Analytics:")
       #Returns to the Segment Analytics Menu
        SegmentAnalytics()


    #Shows the best products by profit and segment
    elif Choice == "3":
        print("Top 10 Products by Segment and Profit:")
        print(Decorating)
        print(Total_Profit_By_Segment_Product.tail(10))
        print(Decorating)
        input("Press enter to go back to Segment Analytics:")
       #Returns to the Segment Analytics Menu
        SegmentAnalytics()

        
    #Shows the worst products by profit and segment
    elif Choice == "4":
        print(Decorating)
        print("Worst 10 Products by Segment and Profit:")
        print(Total_Profit_By_Segment_Product.head(10))
        print(Decorating)
        input("Press enter to go back to Segment Analytics:")
       #Returns to the Segment Analytics Menu
        SegmentAnalytics()

        
    elif Choice == "5":
       #Returns to Main Menu
        MainMenu()
     

    else:
       #Gives Error Message
       print("Error! Please select an option between numbers:1-5 in this sub-menu:")
       input("Press enter to go back to Segment Analytics:")
       #Returns to the Segment Analytics Menu
       SegmentAnalytics()
#============================================================================
def Customer_Analytics():
    print(Decorating)
    print("You are in Customer Analytics SubMenu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Profit based Customer Analytics \n 2: Sales based Customer Analytics")
    print(" 3: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select on option: ")

    if Choice == "1":
        Customer_Analytics_Profit()
        

    elif Choice == "2":
        Customer_Analytics_Sales()
     
    elif Choice == "3":
        MainMenu()
 # error check   
    else:
        print("Invalid input! Please enter a number between 1 and 3.")
        input()
        #input("Press Enter to go back to Customer Analytics:")
    Customer_Analytics()
        
# Profit based Customer Analytics
def Customer_Analytics_Profit():
    print(Decorating)
    print("You are in Profit based Customer Analytics SubMenu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Top 5 Customers that generated the most profit \n 2: Customers that generated the most profit in each region")
    print(" 3: Customers that generated most profit  each year \n 4: Return to Sub Menu: Customer Analytics \n 5: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select an option: ")

# Top 5 Customers that generated the most profit
    if Choice == "1":   
        Profit_columns = SalesData[["Customer Name" , "Profit"]]
        total_Profit = Profit_columns.groupby(by= "Customer Name").sum().sort_values(by = "Profit")
        
        print(Decorating)
        print("\nTop Customers that generated most profit are: " )
        print(total_Profit.tail(5))
        print(Decorating)
        input("Press Enter to go back to Profit based Customer Analytics Sub Menu:")
        Customer_Analytics_Profit()
        
# Customers that generated most profit in Each region
    elif Choice == "2":
        region_profit_subCat = SalesData[["Region", "Profit", "Customer Name"]]
        regions = SalesData.Region.unique()
        print(Decorating)
        print(regions)
        print(Decorating)
        
        for region in regions:
            region_profit = region_profit_subCat.loc[region_profit_subCat["Region"]==region]
            region_total_profit = region_profit.groupby(by="Customer Name").sum().sort_values(by="Profit", ascending = False)
            #region_total_profit = region_total_profit.reset_index()
            print("Customers that generated most profit in " + region + " are: ")
            print(region_total_profit.head(5))
            print(Decorating)
        input("Press Enter to go back to Profit based Customer Analytics Sub Menu:"  )
        Customer_Analytics_Profit()
        
# Customers that generated most profit each year
    elif Choice == "3":
        # List years
        Profit_Year = SalesData
        Profit_Year["Year"] = Profit_Year["Order Date"].dt.year
        Year_List = Profit_Year.Year.unique()
        
        
        # Top 5 customers that generate most profit  
        Customer_Profit = SalesData[["Profit" , "Customer Name" ]]
        Total_Customer_Profit = Customer_Profit.groupby(by= "Customer Name").sum().sort_values(by = "Profit")
        
        
        
        # Top 5 customers that generate most profit in each year
        Year_List_Profit = SalesData[["Customer Name" , "Profit", "Order Date"]]
        
        for Year in Year_List:
            Profit_Year = Year_List_Profit.loc[Year_List_Profit["Order Date"].dt.year==Year]
            Customer_total_profit = Profit_Year.groupby(by= "Customer Name").sum().sort_values(by = "Profit", ascending = False)
            
            
            #print("\033[1;34;47m ")
            print("\nThe top 5 Customers that generate most profit in ", Year, " are: \n")
            #print("\033[1;30;47m ")
            print(Customer_total_profit.head(5))
            print(Decorating)
        input("Press Enter to go back to Profit based Customer Analytics Sub Menu:"  )
        Customer_Analytics_Profit()
        
    elif Choice == "4":
        Customer_Analytics()
    
    elif Choice == "5":
         MainMenu()
         
    else: 
        print("Invalid input! Please enter a number between 1 and 5.")
        
    Customer_Analytics_Profit()
        

# ============================================================================================

# submenu 2 Sales Based customer analytics

def Customer_Analytics_Sales():
    print(Decorating)
    print("You are in Sales based Customer Analytics SubMenu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Top 5 Customers that contributed to most sales \n 2: Customers that contributed to Most sales in each region")
    print(" 3: Customers that contributed to most Sales each year \n 4: Return to Sub Menu: Customer Analytics \n 5: Return to Main Menu" ) 
    print(Decorating)
    Choice = input("Please select an option: ")

# Top 5 Customers that contributed to most sales
    if Choice == "1":   
        sales_columns = SalesData[["Customer Name" , "Sales"]]
        total_sales = sales_columns.groupby(by= "Customer Name").sum().sort_values(by = "Sales", ascending = False )
        
        print(Decorating)
        print("\nTop Customers that contributed to most sales are: " )
        print(total_sales.head(5))
        print(Decorating)
        input("Press Enter to go back to Sales based Customer Analytics Sub Menu:"  )
        Customer_Analytics_Sales()
        
        
# Customers that Contributed to most sales in Each region
    elif Choice == "2":
        region_Sales_Cust = SalesData[["Region", "Sales", "Customer Name"]]
        regions = SalesData.Region.unique()
        print(Decorating)
        print(regions)
        print(Decorating)
        
        for region in regions:
            region_Sales = region_Sales_Cust.loc[region_Sales_Cust["Region"]==region]
            region_total_Sales = region_Sales.groupby(by="Customer Name").sum().sort_values(by="Sales", ascending = False)
            #region_total_profit = region_total_profit.reset_index()
            print("Customers that contributed to most sales in " + region + " are: ")
            print(region_total_Sales.head(5))
            print(Decorating)
        input("Press Enter to go back to Sales based Customer Analytics Sub Menu:"  )
        Customer_Analytics_Sales()
        
            
# Customers that contributed to most sales each year
    elif Choice == "3":
        # List years
        Sales_Year = SalesData
        Sales_Year["Year"] = Sales_Year["Order Date"].dt.year
        Year_List = Sales_Year.Year.unique()
        
        
        # Top 5 customers that contribute to most sales 
        Customer_Sales = SalesData[["Sales" , "Customer Name" ]]
        Total_Customer_Sales = Customer_Sales.groupby(by= "Customer Name").sum().sort_values(by = "Sales")
        
        
        # Top 5 customers that contributed to most sales in each year
        Year_List_Sales = SalesData[["Customer Name" , "Sales", "Order Date"]]
        
        for Year in Year_List:
            Sales_Year = Year_List_Sales.loc[Year_List_Sales["Order Date"].dt.year==Year]
            Customer_total_Sales = Sales_Year.groupby(by= "Customer Name").sum().sort_values(by = "Sales", ascending = False)
            
            #print("\033[1;34;47m ")
            print("\nThe top 5 Customers that contributed to most sales in ", Year, " are: \n")
            #print("\033[1;30;47m ")
            print(Customer_total_Sales.head(5))
            print(Decorating)
        input("Press Enter to go back to Sales based Customer Analytics Sub Menu:"  )
        Customer_Analytics_Sales()
        
            
    elif Choice == "4":
        Customer_Analytics()
    

    elif Choice == "5":
         MainMenu()
         

    else: 
        print("Invalid input! Please enter a number between 1 and 5.")
        
        Customer_Analytics_Sales()
        
# =============================================================================
# Option # 3 Category Analytics Code Begins
#==============================================================================
# Category Analytics (with in Main Menu Option)
def Category_Analytics():
    print(Decorating)
    print("You are in Category Analytics SubMenu \n")
    print("Please choose the following options to proceed:")
    print("\n 1: Profit based category analytics \n 2: Sales based category analytics")
    print(" 3: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select an option: ")
    
    if Choice == "1":
        Category_Analytics_Profit()

    elif Choice == "2":
        Category_Analytics_Sales()
            
    elif Choice == "3":
        MainMenu()

    else:
        print("Invalid input! Please enter a number between 1 and 3.")
        Category_Analytics()

def Category_Analytics_Profit():
    print(Decorating)
    print("You are in Profit based Category Analysis SubMenu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Most profit generated by all Categories \n 2: The 5 states that generated the most profit in each category")
    print(" 3: The 5 states that generated the least profit in each category")
    print(" 4: Most profit generated by Categories each year \n 5: Return to Sub Menu: Category Analytics")
    print(" 6: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select an option: ")

# Most profit generated by all Categories 
    if Choice == "1":     
        Profit_columns = SalesData[["Category" , "Profit"]]
        total_Profit = Profit_columns.groupby(by= "Category").sum().sort_values(by = "Profit")

        print(Decorating)
        print("\nMost Profit generated by categories is: \n" )
        print(total_Profit.head(5))
        print(Decorating)
        
# Pie Chart
        
        # for pie chart for most profit generated by categories
        Cust_Category = SalesData[["Category","Profit"]]
        Customer_Cate_Profit = SalesData.groupby(["Category"]).sum().sort_values(by=["Profit"]).reset_index()
        Categories = list(Customer_Cate_Profit.Category.values)
        TotalProfit = list(Customer_Cate_Profit.Profit.values)
    
     #Gives the colors of pie chart
        symbols = ['cyan','red','yellow']
        plt.figure(figsize = (7,7))
        explode = (0, 0, 0)
        salespiechart = plt.pie(TotalProfit, labels = Categories, explode=explode,shadow= True, colors=symbols,autopct='%1.1f%%')
#plt.axis('equal')
        plt.title("Total Profit Percentage per Category\n")
        plt.show()
        input("Press enter to go back to Profit based category analytics Sub menu")
       #Returns to sub menu
        Category_Analytics_Profit()
# top 5 states that generated Most profit in all Categories
    elif Choice == "2":
        State_profit_Cat = SalesData[["State" , "Profit", "Category"]]
        categories = SalesData.Category.unique()
        print(Decorating)
        print(categories)
        print(Decorating)

        for category in categories:
            State_profit = State_profit_Cat.loc[State_profit_Cat["Category"]==category]
            State_total_profit = State_profit.groupby(by= "State").sum().sort_values(by = "Profit", ascending = False)
         
            print("\nIn " + category + " category the top 5 States in terms of profit are: ")
            print(State_total_profit.head(5))
            print(Decorating)
        input("Press Enter to go back to Profit based category analytics Sub menu")
       #Returns to sub menu
        Category_Analytics_Profit()
        
        
         
# Least profit generating states in all categories
    elif Choice == "3":
        State_profit_Cat = SalesData[["State" , "Profit", "Category"]]
        categories = SalesData.Category.unique()
        print(Decorating)
        print(categories)
        print(Decorating)

        for category in categories:
            State_profit = State_profit_Cat.loc[State_profit_Cat["Category"]==category]
            State_total_profit = State_profit.groupby(by= "State").sum().sort_values(by = "Profit")
         
            print("\nIn " + category + " category the 5 States that generates least profit are : ")
            print(State_total_profit.tail(5))
            print(Decorating)
        input("Press Enter to go back to Profit based category analytics Sub menu")
       #Returns to sub menu
        Category_Analytics_Profit()
        
        
# Most profit generated by each category each year
    elif Choice == "4":
        # List years
        Profit_Year = SalesData
        Profit_Year["Year"] = Profit_Year["Order Date"].dt.year
        Year_List = Profit_Year.Year.unique()
    
    
    # profit generated by categories 
        Category_Profit = SalesData[["Profit" , "Category" ]]
        Total_Category_Profit = Category_Profit.groupby(by= "Category").sum().sort_values(by = "Profit")
    
    
        Year_List_Profit = SalesData[["Category" , "Profit", "Order Date"]]
    
        for Year in Year_List:
            Profit_Year = Year_List_Profit.loc[Year_List_Profit["Order Date"].dt.year==Year]
            Category_total_profit = Profit_Year.groupby(by= "Category").sum().sort_values(by = "Profit")
        
            print("\nMost profit generated by categories in ", Year, " are: \n")
            print(Category_total_profit.head(5))
            print(Decorating)
        input("Press Enter to go back to Profit based category analytics Sub menu")
       #Returns to sub menu
        
        Category_Analytics_Profit()
        
        
    elif Choice == "5":
        Category_Analytics()

    elif Choice == "6":
        MainMenu()

    else: 
        print("Invalid input! Please enter a number between 1 and 6.")
        Category_Analytics_Profit()
        
#==============================================================================
        
# submenu 2 Sales Based category analytics
def Category_Analytics_Sales():
    print(Decorating)
    print("You are in Sales based Category Analytics SubMenu \n")
    print("Please choose following options to proceed:")
    print("\n 1: Most Sales contributed by all Categories \n 2: The 5 states that contributed to most sales in each category")
    print(" 3: The 5 states that contributed to least sales in each category")
    print(" 4: Most Sales by Categories Each year \n 5: Return to Sub Menu: Category Analytics \n 6: Return to Main Menu" )
     
    print(Decorating)
    Choice = input("Please select an option: ")

# Most sales contributed by all Categories 
    if Choice == "1":     
        sales_columns = SalesData[["Category" , "Sales"]]
        total_sales = sales_columns.groupby(by= "Category").sum().sort_values(by = "Sales")

        print(Decorating)
        print("\nMost sales contributed by categories are: " )
        print(total_sales.head(5))
        print(Decorating)
        

        # pie chart
        Cust_Category = SalesData[["Category","Sales"]]
        Customer_Cate_Sales = SalesData.groupby(["Category"]).sum().sort_values(by=["Sales"]).reset_index()
        Categories = list(Customer_Cate_Sales.Category.values)
        TotalSales = list(Customer_Cate_Sales.Sales.values)
    
     #Gives the colors of pie chart
        symbols = ['cyan','red','yellow']
        plt.figure(figsize = (7,7))
        explode = (0, 0, 0)
        salespiechart = plt.pie(TotalSales, labels = Categories, explode=explode,shadow= True, colors=symbols,autopct='%1.1f%%')
        #plt.axis('equal')
        plt.title("Total Sales Percentage per Category\n")
        plt.show()
        input("Press Enter to go back to Sales based category analytics Sub menu")
        Category_Analytics_Sales()
        
# In the category the 5 States that generates most Sales :        
    elif Choice == "2":
        pcode_sale_Cat = SalesData[["State" , "Sales", "Category"]]
        categories = SalesData.Category.unique()
        print(Decorating)
        print(categories)
        print(Decorating)
        
        for category in categories:
            pcode_sale = pcode_sale_Cat.loc[pcode_sale_Cat["Category"]==category]
            pcode_total_sale = pcode_sale.groupby(by= "State").sum().sort_values(by = "Sales")
            
            print("In " + category + " category the top 5 States in terms of generating most sales are: ")
            print(pcode_total_sale.tail(5))
            print(Decorating)
        input("Press Enter to go back to Sales based category analytics Sub menu")
        Category_Analytics_Sales()
        
#  In the category the 5 States that generates Least Sales :           
    elif Choice == "3":
        pcode_sale_Cat = SalesData[["State" , "Sales", "Category"]]
        categories = SalesData.Category.unique()
        print(Decorating)
        print(categories)
        print(Decorating)
        
        for category in categories:
            pcode_sale = pcode_sale_Cat.loc[pcode_sale_Cat["Category"]==category]
            pcode_total_sale = pcode_sale.groupby(by= "State").sum().sort_values(by = "Sales")
            
            print("In " + category + " category the 5 States in terms of generating least sales are: ")
            print(pcode_total_sale.head(5))
            print(Decorating)
        input("Press Enter to go back to Sales based category analytics Sub menu")
        Category_Analytics_Sales()
        
        
# Most sales by each Category each year          
    elif Choice == "4":
         # List years
        Sales_Year = SalesData
        Sales_Year["Year"] = Sales_Year["Order Date"].dt.year
        Year_List = Sales_Year.Year.unique()
        
        
        # Sales contributed by categories 
        Category_Sales = SalesData[["Sales" , "Category" ]]
        Total_Category_Sales = Category_Sales.groupby(by= "Category").sum().sort_values(by = "Sales")
        
        
        Year_List_Sales = SalesData[["Category" , "Sales", "Order Date"]]
        
        for Year in Year_List:
            Sales_Year = Year_List_Sales.loc[Year_List_Sales["Order Date"].dt.year==Year]
            Category_total_Sales = Sales_Year.groupby(by= "Category").sum().sort_values(by = "Sales")
            
            print("\nMost Sales contributed by categories in ", Year, " are: \n")
            print(Category_total_Sales.head(5))
            print(Decorating)
            #Category_Analytics_Sales()
        input("Press Enter to go back to Sales based category analytics Sub menu")
        Category_Analytics_Sales()
        
    elif Choice == "5":
        Category_Analytics()
         

    elif Choice == "6":
        MainMenu()
         
    else:
        print("Invalid input! Please enter a number between 1 and 6.")
        Category_Analytics_Sales()  

        
            
       
       
       
       
#=============================================================================       
def MainMenu():
    print(Decorating)
    print('Welcome to the Office Solutions.\n')
    print(Decorating)
    print(" Main Menu: \n Please Choose the following Inight Menus from following options: \n\n 1. Register New User  \n 2. Customer Analytics \n 3. Category Analysis \n 4. Sub-Category Analytics \n 5. Segment Analytics \n 6. Discount Analytics \n 7. State Analytics \n 8. Logout" )
    Choice = input("Please select from following options: ")
    print(Decorating)
    
    while True:
        break
        print(Choice)
#         
#     
    if Choice == "1":
        register ()

    elif Choice == "2":
        Customer_Analytics()
        
# Customer Analytics        

    elif Choice == "3":
        Category_Analytics()
        
# # Category Analytics       

    elif Choice == "4": 
#         
        SubCategoryAnalytics()

    elif Choice == "5":
        SegmentAnalytics()
# # Discount Analytics
    elif Choice == "6":
        DiscountAnalytics()
# # State Analytics
    elif Choice == "7":
        StateAnalytics()
        
    elif Choice == "8":
        print("You are logged off. Thank you for using Office Solutions work app")
        SystemExit()
        login()
        
# =============================================================================
        
# =============================================================================
    else:
         print("\nInvalid Input. Please enter correct number from 1 through 8.")
         print("\n")
         MainMenu()
#         login()
# =============================================================================
# =============================================================================
# LOg in code begins
#==============================================================================
def login():
    conn = sqlite3.connect('OS_Employee.db')
    with conn:
        loginTest = False
        while not loginTest:
            cur = conn.cursor()
            try:
                print(Decorating)
                print("Welcome to the Office Solutions App!")
                print(Decorating)
                userEmail = input("Please enter your Email: ")
                userEmail = userEmail.strip()
# Checking for Blank UserEmail
                while not userEmail:
                    userEmail = input("Email cannot be blank. Please enter your Email: ")
                    userEmail = userEmail.strip()
                
# User Password required                
                userPassword = input("Please enter your password: ")
                userPassword = userPassword.strip()
 # checking for Blank Password
                while not userPassword:
                    userPassword = input("Password cannot be blank. Please enter your password: ")
                    userPassword = userPassword.strip()
                    
                cur.execute("SELECT COUNT (*) FROM Employee WHERE(Email = '" + userEmail.lower() +"' AND Password = '" + userPassword + "') ")
                results = cur.fetchone()
                (results[0])
                if results[0]==1:

                    print(" \nLogin successful!")
                    loginTest = True
                    postLogin()
                else:
                    print("\nLogin unsuccessful! Please try logging in again.")
            except sqlite3.Error as e:
                print(e)
    conn.close()
#-------------------------------------------------------------------------------------------------------------------
# User can register for new user or go back to main menu
def postLogin ():
    print("==============================================")
    print ("1. Update Password. \n2. Main Menu.")

 # checking for correct input between option 1-2    
    ans = 0
    while not ans:
        try:
            ans = int(input('Please choose an option:'))
            if ans not in (1,2):
                raise ValueError
        except ValueError:
                ans = 0
                print ('\n- Note: Enter a number corresponding to the menu. ')
    
    if ans == 1:
         return PasswordChange()
    elif ans == 2:
         return MainMenu()
  
#-------------------------------------------------------------------------------------------------------------------
    

#-------------------------------------------------------------------------------------------------------------------
# Registeration of new user 
def register ():
    newID = input('Please enter an Employee ID for the new user: ')
# Checking for Blank input from user for Employee ID. Employee ID should not be Blank. 
    while not newID:
        newID = input("Employee ID cannot be blank. Please enter Employee ID: ")
        newID = newID.strip()
    
    conn = sqlite3.connect("OS_Employee.db")
    with conn:
        cur = conn.cursor()
        try:
            cur.execute("SELECT COUNT(EmployeeID) FROM Employee WHERE EmployeeID = '" + newID + "'")
            results = cur.fetchone()
 # Checking for existing EmployeeId in the system.           
            while results[0] != 0:
                print("\nID already exists! Please choose another.\n")
                newID = input('Enter new UserID: ')
                
                cur.execute("SELECT COUNT(EmployeeID) FROM Employee WHERE EmployeeID = '" + newID + "'")
                results = cur.fetchone()
#-------------------------------------------------------------------------------------------------------------------
# Take user input for first name
            firstName = input('Please enter first name: ')
            firstName = (firstName.title())
 # Checking for blank first name from the user
            while not firstName:
                firstName = input('First Name cannot be blank. Please enter your first name: ')        
                firstName = (firstName.title())
#-------------------------------------------------------------------------------------------------------------------
# Take user input for the last name
            lastName = input('Please enter last name: ')
            lastName = (lastName.title())
# Checking for blank last name from the user
            while not lastName:
                lastName = input('Last Name cannot be blank. Please enter your last name: ')        
                lastName = (lastName.title())
#Email-------------------------------------------------------------------------------------------------------------------
# take user input for the emial address 
            email = input('Please enter an email address for new user: ')           
            email = email.strip()
            email = email.lower()
# Checking for blanks in email input by the user
            while not email: 
                email = input('email cannot be blank. Please enter your email: ')
                email = email.strip()
                email = email.lower()
#-------------------------------------------------------------------------------------------------------------------
# Take user input for Password
            password = input('Please enter a password for the new User: ')
            password = password.lower()
 # checking for blanks in password entry by the user
            while not password: 
                password = input('password cannot be blank. Please enter your password: ')
                password = password.lower()
            
            cur.execute("SELECT COUNT(Email) FROM Employee WHERE Email = '" + email + "'")
            results = cur.fetchone()
#------------------------------------------------------------------------------------------------------------------------------------- 
# Checking for existing Email address in the databse
            while results[0] != 0:
                print("\nThis email already exists! Please choose another.\n")
                email = input("Please enter an email address for new user: ")           
                email = email.strip()
                email_val = False

                while not email_val: 
                    if '@' in email:
                            if email.count('@') == 1: 
                                if '.' in email:
                                    email_val = True
                                    email = email.lower()
                                else:
                                    print("\n- Note: Email address must contain one '@' symbol and at least one '.' symbol.")
                                    email = input("Please re-enter email address: ")
                                    email = email.strip()
                                    email = email.lower()
                            else:
                                 print("\n- Note: Email address must contain one '@' symbol and at least one '.' symbol.")
                                 email = input("Please re-enter email address: ")
                                 email = email.strip()
                                 email = email.lower()
                    else:
                        print("\n- Note: Email address must contain one '@' symbol and at least one '.' symbol.")
                        email = input("Please re-enter email address: ")
                        email = email.strip() 
                        email = email.lower()
                cur.execute("SELECT COUNT(Email) FROM Employee WHERE Email = '" + email + "'")
                results = cur.fetchone()
                
#Password for Email Exists-------------------------------------------------------------------------------------------------------------------             
                password = input('Please enter a password for the new User: ')
                while len(password) == 0 or (not password.isalnum()): #Remove 'len(lastName) <= 1 or' if you dont want to limit name
                    print ('\n- Note: Password must contain only letters and numbers')
                    password = input('Please re-enter password: ')       
                    password = (password.lower())
            
            cur = conn.cursor()
            
            cur.execute("INSERT INTO Employee (EmployeeID, FirstName, LastName, Email, Password) VALUES ('" + newID + "', '" + firstName + "', '" + lastName + "', '" + email + "', '" + password + "')")
            
            conn.commit()
            cur.execute("SELECT COUNT (*) FROM Employee WHERE EmployeeID = '" + newID + "'")
                        
            results = cur.fetchone()
          
            if results[0] == 1:
                print("\n")
                print("New user has been added successfully!")
                          
                while True:
                    print(Decorating)
                    print ("Welcome! Please choose an option:")
                    print ("1. Login as new user")
                    print ("2. Go to Main menu")
                    print(Decorating)
                    ans=input("What would you like to do? ")
                    if ans == "1":
                        login()
                    elif ans == "2":
                        print("2")
                        MainMenu()
                    elif ans =="":
                        print('Please choose from the following: ')
              
            else:
                print("Registration unsuccessful!")
            
                    
        except sqlite3.Error as e:
            print(e)

#Changes Password when user log
            
def PasswordChange():
    EmployeeId = input('Please enter an Employee ID number: ')
    # Checking for blank  employee ID from the user
    while not EmployeeId:
                        EmployeeId = input("Employee ID field cannot be blank. Please enter Employee ID: ")
    CurrentPW = input('Please enter current password: ')
    #Checking for blank current password from the user
    while not CurrentPW:
                  CurrentPW = input('Current Password field cannot be blank. Please enter your current password: ')      
    NewPW = input('Please enter the new password: ')
    # Checking for blank new password from the user
    while not NewPW:
                  NewPW = input('New Password field cannot be blank. Please enter your new password: ')
    # Checking for existing EmployeeId and password in the system.           
    cur = conn.cursor()
    cur.execute("SELECT COUNT(Password),COUNT(EmployeeID) FROM Employee WHERE EmployeeID = '" + EmployeeId + "'"" AND Password = '" + CurrentPW + "'")    
    results = cur.fetchone()
    while not results[0] != 0:
        print("\nID number and current password do not exist! Please enter correct ID number and correct current password.\n")
        EmployeeId = input('Enter correct Employee ID number: ')
        CurrentPW = input('Please enter current password: ')
                
        cur.execute("SELECT COUNT(Password),COUNT(EmployeeID) FROM Employee WHERE EmployeeID = '" + EmployeeId + "'"" AND Password = '" + CurrentPW + "'")
        results = cur.fetchone()
    
    
#Updates a users password
    try:
        with conn:
            cur = conn.cursor()
            ChangedValues = "UPDATE Employee SET Password = ('{}') WHERE (EmployeeID = ('{}') AND Password = ('{}'))"
            ChangedString = ChangedValues.format(NewPW, EmployeeId, CurrentPW)
            cur.execute(ChangedString)
            conn.commit
            print(ChangedString)
            conn.commit()
            cur.execute("SELECT * FROM Employee WHERE EmployeeID = ('{}')".format(EmployeeId))
            results = cur.fetchone()
            print(results)
            if results[4] == NewPW:
                print("\n")
                print("Password changed successfully!")
            input("Press Enter to continue..")
            MainMenu()
            
    except Exception as e:
        print("Failed connection" + str(e))
    
    conn.close()
    

             
# Checking for Blank input from user for Employee ID. Employee ID should not be Blank. 
login()        
        
        
        
        
   
# =============================================================================
# Option # 2 Customer Analytics Code Begins
#==============================================================================
# Customer Analytics (with in Main Menu Option)

        



#========================================================================================