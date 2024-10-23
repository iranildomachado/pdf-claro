import tabula
import tablib
from os import walk
from pathlib import Path


class PDF_CLARO:

    def __init__(self) -> None:
        self.pdf_paths_list = list()
        self.__open_files()


    def __open_files(self) -> list:
        """ Valida e abre o arquivo em pdf. """

        try:

            for path, diretorios, arquivos in walk("pdf"):
                for arquivo in arquivos:
                    pdf_file = Path(f'pdf/{arquivo}')

                    if not pdf_file.exists():
                        print(f'\n❌ Arquivo não existe ou não localidado.\n')
                        continue

                    else:
                        self.pdf_paths_list.append(pdf_file)

        except Exception as exc:
            print(f'\n❌ EXCEPTION IN OPEN_FILE => {exc}\n')
            return None


    def get_data_pdf(self):
        """ Obtem as informações do PDF necessárias. """

        try:
            pdf_dataframe = list()
            for path in self.pdf_paths_list:

                with open(path, 'rb') as pdf_file:
                    pdf_reader = tabula.read_pdf(pdf_file, pages="all")
                    pdf_tabula = pdf_reader[0]
                    pdf_dataframe.append(pdf_tabula)

            return pdf_dataframe

        except Exception as exc:
            print(f'\n❌ EXCEPTION IN get_data_pdf => {exc}\n')
            return None


    def export_xlsx(self):
        """ Exporta os dados da tabela para excel. """

        try:
            dataset = tablib.Dataset()
            dataset.headers = ["Técnico", "RG", "Data Agendamento", "Hora Agendamento Incio", "Hora Agendamento Fim", "Hora Entrada", "Hora Saída"]

            set_dataframe = self.get_data_pdf()
            # print(set_dataframe[0])

            for data in set_dataframe:
                dataset.append(
                    (
                        data.iloc[21][1],
                        str(data.iloc[21][2]).split("Login")[0].replace("RG", "").strip(),
                        str(data.iloc[26][1]).split(" ")[0].strip(),
                        str(data.iloc[26][1]).split(" ")[1].split("-")[0],
                        str(data.iloc[26][1]).split("-")[1].strip(),
                        str(data.iloc[26][2]).split("Hora")[1].replace("Entrada", "").strip(),
                        str(data.iloc[26][2]).split("Hora")[2].replace("Saida", "").strip()
                    )
                )

            # exporting to excel file
            with open('report_claro.xlsx', 'wb') as f:
                f.write(dataset.export('xlsx'))


        except Exception as exc:
            print(f'\n❌ EXCEPTION IN export_xlsx => {exc}\n')
            return None            

        
PDF_CLARO().export_xlsx()