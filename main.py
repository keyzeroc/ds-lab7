import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt

# Читання файлу Excel
file_path = "Pr_1.xls"
data_sales = pd.read_excel(file_path, sheet_name="Лист1")
data_regions = pd.read_excel(file_path, sheet_name="Лист2")

print("Дані з першої таблиці (продажі):")
print(data_sales.head())

print("Дані з другої таблиці (регіони):")
print(data_regions.head())


# Аналіз основних характеристик даних
print("Опис даних продажів:")
print(data_sales.describe())

print("Опис даних регіонів:")
print(data_regions.describe())

# Перевірка наявності пропущених значень
print("Пропущені значення у даних продажів:")
print(data_sales.isnull().sum())

print("Пропущені значення у даних регіонів:")
print(data_regions.isnull().sum())


# Обчислення продажів та прибутку
data_sales["Продажі"] = (
    data_sales["КільКість реалізацій"] * data_sales["Ціна реализації"]
)
data_sales["Прибуток"] = data_sales["Продажі"] - (
    data_sales["КільКість реалізацій"] * data_sales["Собівартість одиниці"]
)

print("Дані з доданими стовпчиками 'Продажі' та 'Прибуток':")
print(data_sales.head())


# Підготовка даних для МНК
X = data_sales[["КільКість реалізацій", "Собівартість одиниці", "Ціна реализації"]]
y = data_sales["Прибуток"]

# Додавання константи до матриці ознак
X = sm.add_constant(X)

# Підгонка моделі
model = sm.OLS(y, X).fit()

print("Результати моделі МНК:")
print(model.summary())


# Об'єднання даних продажів та регіонів
merged_data = pd.merge(data_sales, data_regions, on="Код магазину")

# Групування за регіонами і місяцями
grouped_data = (
    merged_data.groupby(["Регіон", "Місяц"]).agg({"Прибуток": "sum"}).reset_index()
)

# Створення таблиці та графіка
pivot_table = grouped_data.pivot(index="Місяц", columns="Регіон", values="Прибуток")
pivot_table.plot(kind="bar", figsize=(12, 8))

plt.title("Динаміка зміни прибутку за регіонами")
plt.xlabel("Місяць")
plt.ylabel("Прибуток")
plt.legend(title="Регіон")
plt.show()

# Збереження таблиці у файл
pivot_table.to_excel("profit_forecast_by_region.xlsx")
