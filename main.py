from connect_api7 import ConnectApi7
import pandas as pd
import openpyxl, os

class Config():
    #необходустановить полный путь к папке с файлами
    dir = "C:/Users/maks-/OneDrive/Рабочий стол/test_app/"

if __name__ == "__main__":

    paths = list()

    for dirpath, dirnames, filenames in os.walk(Config.dir):
        for filename in filenames:
            paths.append(os.path.join(dirpath, filename))

    api7 = ConnectApi7()
    documents = []

    for path in paths:
        document = []
        if path[-4:] == ".m3d":
            api7.openDocument(path)

            for variable in api7.getVariables3D():
                document.append(variable.Value)
            document.append(path)

            api7.closeDocument()
            documents.append(document)

        if len(document) > 1:
            dataFrame = pd.DataFrame(documents, columns=['D','d','h','g','lg','path'])
            writer = pd.ExcelWriter(Config.dir + 'parametrsDetails.xlsx', engine='xlsxwriter')
            dataFrame.to_excel(writer, 'Параметры детали - "Шайба"')
            writer.close()