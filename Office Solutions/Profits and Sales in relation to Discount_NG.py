# -*- coding: utf-8 -*-
"""
Created on Thu Nov 21 09:59:09 2019

@author: nikhi
"""
#Importing from different libraries:matplotlib,seaborn, and scipy. Providing aliases
import matplotlib.pyplot as plt
import seaborn as sns
import scipy.stats 


import pandas as pd

xl = pd.ExcelFile("SalesData.xlsx")
SalesData = xl.parse("Orders")
dashes = "="*25


Profits_by_Discount = SalesData.groupby(["Discount"]).agg({'Profit':sum})
print(Profits_by_Discount)    
 
def Subcategories_BY_Discount():
    Subcategories_BY_Discount = SalesData[["Sub-Category","Discount"]]
    Subcategory_Max_Rate = Subcategories_BY_Discount.groupby(by="Sub-Category").max()
    print(dashes)
    print(Subcategory_Max_Rate)
    print(dashes)
    
    Subcategory_Min_Rate = Subcategories_BY_Discount.groupby(by="Sub-Category").min()
    print(dashes)
    print(Subcategory_Min_Rate)
    print(dashes)
Subcategories_BY_Discount()
print()

def relationship_between_sales_discount():
    sns.regplot(x = "Discount", y = "Sales", fit_reg = True, data = SalesData)
    plt.xlabel("Discount Rate")
    plt.ylabel("Sales")
    plt.title("Relationship between Sales and Discount \n")
    plt.show()
relationship_between_sales_discount()
print()
def relationship_between_profit_discount():
    sns.regplot(x = "Discount", y = "Profit", fit_reg = True, data = SalesData)
    plt.xlabel("Discount Rate")
    plt.ylabel("Profit")
    plt.title("Relationship between Profit and Discount \n")
    plt.show()
relationship_between_profit_discount()
print()

def Orders_By_Discount_Number():
    SalesData["Discount"] = pd.Categorical(SalesData["Discount"])
    sns.countplot(x="Discount", data=SalesData)
    plt.xlabel("Discount Rate Offered")
    plt.show()
Orders_By_Discount_Number()
print()
    
