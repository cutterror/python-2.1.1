from dataset import DataSet
from report import Report

file_name = input("Введите название файла: ")
profession_name = input("Введите название профессии: ")
data_set = DataSet(file_name, profession_name)
data_set.statistic.print_statistics()
report = Report(data_set.statistic)
report.generate_excel()
