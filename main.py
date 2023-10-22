import sys

import sqlalchemy
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMainWindow, QApplication, QPushButton,QFileDialog,QMessageBox
import pandas as pd
from PyQt6 import QtWidgets, QtGui
from PyQt6.QtGui import QColor
from sqlalchemy import create_engine
import datetime
from dbsettings import database_parametres
from interface import Ui_MainWindow
from interface import Ui_Dialog
import psycopg2
from find_file_origin import update_excel_file
from parsers_two_table.entire_parser import parse_articles_to_excel
from parsers_two_table.findPeopleWithoutID import update_author_id
from parsers_two_table.findOrganisationsWithoutID import update_org_id
from find_duplicates import deduplicate_excel
from interface import Ui_Dialog2
from find_new_elibrary_id import update_elibrary_id


class MyDialog(QtWidgets.QDialog):
    def __init__(self, data_1):
        super(MyDialog, self).__init__()
        self.ui = Ui_Dialog2()
        self.ui.setupUi(self)
        self.fillDialogTables(data_1)
        self.ui.pushButton.clicked.connect(lambda: self.fill_upper_table(data_1))
        self.ui.pushButton_28.clicked.connect(lambda: self.fill_lower_table(data_1))

    def fillDialogTables(self, data_1):
        self.ui.tableWidget_44.clearContents()
        self.ui.tableWidget_228.clearContents()
        for i in range(1):
            for j in range(4):
                item = QtWidgets.QTableWidgetItem(str(data_1[i][j]))
                self.ui.tableWidget_228.setItem(i, j, item)
        for i in range(1):
            for j in range(4, 8):
                item = QtWidgets.QTableWidgetItem(str(data_1[i][j]))
                self.ui.tableWidget_44.setItem(i, j - 4, item)
    def fill_upper_table(self, data_1):
        try:
            conn = psycopg2.connect(database=database_parametres['dbname'],
                                    user=database_parametres['user'],
                                    password=database_parametres['password'],
                                    host=database_parametres['host'],
                                    port=database_parametres['port'])
            cursor = conn.cursor()
            data_list = []
            data2_list = []
            num_columns = self.ui.tableWidget_44.columnCount()
            row_index = 0
            for col_index in range(num_columns):
                item = self.ui.tableWidget_44.item(row_index, col_index)
                item_2 = self.ui.tableWidget_228.item(row_index, col_index)
                if item is not None:
                    data = item.text()
                    data_list.append(data)
                    data2 = item_2.text()
                    data2_list.append(data2)
            print(data_list)
            update_query = f"UPDATE authors_organisations SET counter= %s, author_id =%s, author_name = %s, author_initials = %s WHERE counter = %s"
            cursor.execute(update_query, (data_list[0], data_list[1], data_list[2], data_list[3], data2_list[0]))
            conn.commit()
            data_1.pop(0)
            if len(data_1) > 0:
                self.fillDialogTables(data_1)
            else:
                self.close()
        except Exception as e:
            print(f"An error occurred: {e}")

    def fill_lower_table(self, data_1):
        try:
            data_1.pop(0)
            if len(data_1) > 0:
                self.fillDialogTables(data_1)
            else:
                self.close()
                sys.exit()
        except Exception as e:
            # Handle the exception here
            print(f"An error occurred: {e}")
class Dialog(QtWidgets.QDialog):
    def __init__(self, data_1, data_2, index_array_1, index_array_2):
        super(Dialog, self).__init__()
        self.ui_dialog = Ui_Dialog()
        self.ui_dialog.setupUi(self)
        self.fillDialogTables(data_1, data_2, index_array_1, index_array_2)
        self.button_lower_table_choice = self.findChild(QPushButton, "button_lower_table_choice")
        self.button_lower_table_choice.clicked.connect(lambda: self.lowertableClicked(data_1, data_2, index_array_1, index_array_2))
        self.button_upper_table_choice = self.findChild(QPushButton, "button_upper_table_choice")
        self.button_upper_table_choice.clicked.connect(lambda: self.uppertableClicked(data_1, data_2, index_array_1, index_array_2))
        self.ui_dialog.label.setText(f'Осталось выбрать {len(data_1)} строк')

    def fillDialogTables(self, data_1, data_2, index_array_1, index_array_2):
        self.ui_dialog.row_from_database.clearContents()
        self.ui_dialog.row_from_excel.clearContents()
        some_row_index = 0
        some_column_index = 0
        for i in range(len(data_1)):
            for j in range(26):
                item = QtWidgets.QTableWidgetItem(str(data_1[i][j]))
                self.ui_dialog.row_from_database.setItem(i, j, item)
        for i in range(len(data_2)):
            for j in range(26):
                item = QtWidgets.QTableWidgetItem(str(data_2[i][j]))
                self.ui_dialog.row_from_excel.setItem(i, j, item)
    def fillDialogTables(self, data_1, data_2, index_array_1, index_array_2):
        self.ui_dialog.row_from_database.clearContents()
        self.ui_dialog.row_from_excel.clearContents()
        some_row_index = 0
        some_column_index = 0
        for i in range(len(data_1)):
            for j in range(26):
                item = QtWidgets.QTableWidgetItem(str(data_1[i][j]))
                self.ui_dialog.row_from_database.setItem(i, j, item)
        for i in range(len(data_2)):
            for j in range(26):
                item = QtWidgets.QTableWidgetItem(str(data_2[i][j]))
                self.ui_dialog.row_from_excel.setItem(i, j, item)

    def lowertableClicked(self, data_1, data_2, index_array_1, index_array_2):
        if len(data_1) > 0:
            data_1.pop(0)
            data_2.pop(0)
            index_array_2.pop(0)
            self.ui_dialog.label.setText(f'Осталось выбрать {len(data_1)} строк')
        if len(data_1) > 0:
            self.fillDialogTables(data_1, data_2, index_array_1, index_array_2)
        else:
            self.close()

    def uppertableClicked(self, data_1, data_2, index_array_1, index_array_2):
        if len(data_1) > 0:
            data_1.pop(0)
            data_2.pop(0)
            index_array_1.pop(0)
            self.ui_dialog.label.setText(f'Осталось выбрать {len(data_1)} строк')
        if len(data_1) > 0:
            self.fillDialogTables(data_1, data_2, index_array_1, index_array_2)
        else:
            self.close()

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.widget_onlyicons.hide()
        self.ui.stackedWidget.setCurrentIndex(0)
        self.ui.home_button_iconexpandedwidget.setChecked(True)
        self.import_button_onlyiconwidget = self.findChild(QPushButton, "import_button_onlyiconwidget")
        self.import_button_onlyiconwidget.clicked.connect(self.importButtonClickHandler)
        self.export_button = self.findChild(QPushButton, "pushButton")
        self.export_button.clicked.connect(self.getYearAndSurname)
        self.import_button_expandedwidget = self.findChild(QPushButton, "import_button_expandedwidget")
        self.import_button_expandedwidget.clicked.connect(self.importButtonClickHandler)
        self.search_button = self.findChild(QPushButton, "Primary")
        self.search_button.clicked.connect(self.get_text)
        self.preview_button = self.findChild(QPushButton, "Primary_3")
        self.preview_button.clicked.connect(self.getYearAndSurname)
        # self.search_button = self.findChild(QPushButton, "general_data_export_button")
        # self.search_button.clicked.connect(self.get_test_auf)
        self.user_button = self.findChild(QPushButton, "user_button")
        self.user_button.clicked.connect(lambda: self.authorsReferenceToSQL(database_parametres))
        self.add_one_row_button = self.findChild(QPushButton, "add_one_row_button")
        self.add_one_row_button.clicked.connect(self.addOneRowToDB)
        self.full_search_button = self.findChild(QPushButton, "pushButton_3")
        self.full_search_button.clicked.connect(lambda: self.search(self.ui.tableWidget_article_2, self.ui.textEdit_22, self.next_result_button, self.previous_result_button, self.ui.handleItemChanged))
        self.full_search_button_2 = self.findChild(QPushButton, "pushButton_7")
        self.full_search_button_2.clicked.connect(lambda: self.search(self.ui.tableWidget_article_author, self.ui.textEdit_3, self.next_result_button_2, self.previous_result_button_2, self.ui.handleItemChanged_2))
        self.full_search_button_3 = self.findChild(QPushButton, "pushButton_10")
        self.full_search_button_3.clicked.connect(lambda: self.search(self.ui.tableWidget_authors, self.ui.textEdit_4, self.next_result_button_3, self.previous_result_button_3, self.ui.handleItemChanged_3))

        self.next_result_button = self.findChild(QPushButton, "pushButton_4")
        self.next_result_button.clicked.connect(lambda: self.scroll_to_next_result(self.ui.tableWidget_article_2, self.ui.handleItemChanged))
        self.next_result_button.setEnabled(False)

        self.previous_result_button = self.findChild(QPushButton, "pushButton_6")
        self.previous_result_button.clicked.connect(lambda: self.scroll_to_previous_result(self.ui.tableWidget_article_2, self.ui.handleItemChanged))
        self.previous_result_button.setEnabled(False)

        self.next_result_button_2 = self.findChild(QPushButton, "pushButton_9")
        self.next_result_button_2.clicked.connect(lambda: self.scroll_to_next_result(self.ui.tableWidget_article_author, self.ui.handleItemChanged_2))
        self.next_result_button_2.setEnabled(False)

        self.previous_result_button_2 = self.findChild(QPushButton, "pushButton_8")
        self.previous_result_button_2.clicked.connect(
            lambda: self.scroll_to_previous_result(self.ui.tableWidget_article_author, self.ui.handleItemChanged_2))
        self.previous_result_button_2.setEnabled(False)

        self.next_result_button_3 = self.findChild(QPushButton, "pushButton_11")
        self.next_result_button_3.clicked.connect(
            lambda: self.scroll_to_next_result(self.ui.tableWidget_authors, self.ui.handleItemChanged_3))
        self.next_result_button_3.setEnabled(False)

        self.previous_result_button_3 = self.findChild(QPushButton, "pushButton_12")
        self.previous_result_button_3.clicked.connect(
            lambda: self.scroll_to_previous_result(self.ui.tableWidget_authors, self.ui.handleItemChanged_3))
        self.previous_result_button_3.setEnabled(False)

        self.search_results = []
        self.current_result_index = 0
        self.signal_connected = True

    def showDialog(self, data_1, data_2, index_array_1, index_array_2):
        self.dialog_instance = Dialog(data_1, data_2, index_array_1, index_array_2)
        self.dialog_instance.show()
        self.dialog_instance.exec()
    def showDialog_2(self, data_1):
        try:
            dialog_instance = MyDialog(data_1)
            dialog_instance.show()
            dialog_instance.exec()
        except Exception as e:
            print(f"An error occurred: {str(e)}")

    # def deleteRowOnChoice(self, data_2):
    #     item_id = data_2[0][0]
    #     author_id = data_2[0][1]
    #     author_name = data_2[0][2]
    #     print(item_id, author_id, author_name)
    #     conn = psycopg2.connect(database=database_parametres['dbname'],
    #                             user=database_parametres['user'],
    #                             password=database_parametres['password'],
    #                             host=database_parametres['host'],
    #                             port=database_parametres['port'])
    #     cur = conn.cursor()
    #     try:
    #         delete_query = """
    #                DELETE FROM article_author
    #                 WHERE item_id = %s AND
    #                 author_id = %s AND
    #                 author_name = %s
    #                """
    #         cur.execute(delete_query, (item_id, author_id, author_name))
    #         conn.commit()
    #         print("Row deleted successfully.")
    #
    #     except Exception as e:
    #         print("Error:", e)
    #         conn.rollback()
    #
    #     finally:
    #         cur.close()
    #         conn.close()


    def search(self, table_widget, textEdit, next_btn, prev_btn, handleItemChanged):
        try:
            if self.signal_connected:
                table_widget.itemChanged.disconnect(handleItemChanged)
                self.signal_connected = False
            self.clear_highlighting(table_widget)
            search_text = textEdit.toPlainText().strip()
            self.clear_search_results(table_widget, next_btn, prev_btn)
            for row in range(table_widget.rowCount()):
                for column in range(table_widget.columnCount()):
                    item = table_widget.item(row, column)
                    if item and item.text() == search_text:
                        self.search_results.append((row, column))
            if self.search_results:
                self.current_result_index = 0
                self.highlight_current_result(table_widget)
                self.scroll_to_current_result(table_widget)
                next_btn.setEnabled(len(self.search_results) > 1)
                prev_btn.setEnabled(len(self.search_results) > 1)
            if not self.signal_connected:
                table_widget.itemChanged.connect(handleItemChanged)
                self.signal_connected = True
        except Exception as e:
            print(f"An error occurred: {str(e)}")

    def clear_search_results(self, table_widget, next_btn, prev_btn):
        for row in range(table_widget.rowCount()):
            item = table_widget.item(row, 0)
            if item is not None:
                item.setBackground(Qt.GlobalColor.transparent)
        self.search_results = []
        self.current_result_index = 0
        next_btn.setEnabled(False)
        prev_btn.setEnabled(False)

    def scroll_to_current_result(self, table_widget):
        if self.search_results:
            row, column = self.search_results[self.current_result_index]
            item = table_widget.item(row, column)
            if item:
                table_widget.scrollToItem(item)


    def highlight_current_result(self, table_widget):
        if self.search_results:
            row, _ = self.search_results[self.current_result_index]
            for column in range(table_widget.columnCount()):
                table_widget.item(row, column).setBackground(QColor(238, 221, 102))

    def scroll_to_next_result(self, table_widget, handleItemChanged):
        try:
            if self.signal_connected:
                table_widget.itemChanged.disconnect(handleItemChanged)
                self.signal_connected = False
            if self.search_results:
                self.current_result_index = (self.current_result_index + 1) % len(self.search_results)
                self.clear_highlighting(table_widget)
                self.highlight_current_result(table_widget)
                self.scroll_to_current_result(table_widget)
            if not self.signal_connected:
                table_widget.itemChanged.connect(handleItemChanged)
                self.signal_connected = True
        except Exception as e:
            print(f"An error occurred: {e}")

    def scroll_to_previous_result(self, table_widget, handleItemChanged):
        if self.signal_connected:
            table_widget.itemChanged.disconnect(handleItemChanged)
            self.signal_connected = False
        if self.search_results:
            self.current_result_index = (self.current_result_index - 1) % len(self.search_results)
            self.clear_highlighting(table_widget)
            self.highlight_current_result(table_widget)
            self.scroll_to_current_result(table_widget)
        if not self.signal_connected:
            table_widget.itemChanged.connect(handleItemChanged)
            self.signal_connected = True

    def clear_highlighting(self, table_widget):
        for row in range(table_widget.rowCount()):
            for column in range(table_widget.columnCount()):
                item = table_widget.item(row, column)
                if item is not None:
                    item.setBackground(Qt.GlobalColor.transparent)

    def process_data(self, where):
        try:
            connection = psycopg2.connect(
                dbname=database_parametres['dbname'],
                user=database_parametres['user'],
                password=database_parametres['password'],
                host=database_parametres['host'],
                port=database_parametres['port']
            )
            cursor = connection.cursor()
            sql_query = """
            with cte as (SELECT 
			linkurl, count(DISTINCT author_id) as author_count
			FROM
            article 
			JOIN
			article_authors_linkage USING(item_id)
			JOIN 
			authors_organisations USING(counter)
			GROUP BY linkurl)
            SELECT * FROM (SELECT 
			article.linkurl,
            article.doi,
            article.year,
            article.title_article,
            article.publisher,
           	article.type,
            article.risc,
            article.issn,
            article.edn,
            authors_organisations.author_id,
			authors_organisations.author_name,
			authors_organisations.author_initials,
            authors_organisations.org_id,
            authors_organisations.org_name,
			cte.author_count,
			COUNT(author_id) over affilations_cnt as affilations_count
        FROM
            article 
		JOIN
			article_authors_linkage USING(item_id)
		JOIN 
			authors_organisations USING(counter)
		JOIN
			cte USING(linkurl)
		window affilations_cnt as (partition by linkurl, author_id)) AS subquery
		WHERE subquery.org_id = 570
            """
            if where:
                sql_query += f"AND {where[0]}"
            sql_query += "ORDER BY doi"
            cursor.execute(sql_query)
            result = cursor.fetchall()
            columns = [desc[0] for desc in cursor.description]
            df = pd.DataFrame(result, columns=columns)
            cursor.close()
            connection.close()

            excel_template_path = "shablon_kbpr.xlsx"
            df_template = pd.read_excel(excel_template_path)
            df_template['URL'] = df['linkurl']
            df_template['Идентификатор DOI *'] = df['doi']
            df_template['Количество авторов *'] = df['author_count']
            df_template['Фамилия *'] = df['author_name']
            df_template['Имя *'] = df['author_initials']
            # df_template['Отчество'] = df['patronymic']
            # df_template['Должность *'] = df['position']
            # df_template['Ученая степень *'] = df['academic_degree']
            # df_template['Тип трудовых отношений *'] = df['employment_relationship']
            # df_template['Год рождения *'] = df['birth_year']
            df_template['Количество аффиляций *'] = df['affilations_count']
            df_template['Аффиляция *'] = df['org_name']
            df_template['Дата публикации *'] = pd.to_datetime(df['year'], format='%Y').dt.strftime('01/01/%Y')
            df_template['Наименование публикации *'] = df['title_article']
            df_template['Наименование издания *'] = df['publisher']
            df_template['Вид издания  *'] = df['type']
            df_template['Идентификатор РИНЦ'] = df['risc']
            df_template['Идентификатор ISSN'] = df['issn']
            df_template['Идентификатор EDN'] = df['edn']


            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            output_path = f"shablon_kbpr_{timestamp}.xlsx"
            df_template.to_excel(output_path)
            QMessageBox.information(self, "Экспорт", "Excel файл по шаблону кбпр создан!")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
        finally:
            try:
                cursor.close()
            except:
                pass
            try:
                connection.close()
            except:
                pass

    def process_data_advanced(self, where):
            try:
                connection = psycopg2.connect(
                    dbname=database_parametres['dbname'],
                    user=database_parametres['user'],
                    password=database_parametres['password'],
                    host=database_parametres['host'],
                    port=database_parametres['port']
                )

                cursor = connection.cursor()

                sql_query = """
             with cte as (SELECT 
			linkurl, count(DISTINCT author_id) as author_count
			FROM
            article 
			JOIN
			article_authors_linkage USING(item_id)
			JOIN 
			authors_organisations USING(counter)
			GROUP BY linkurl)
            SELECT * FROM (SELECT 
			article.linkurl,
            article.doi,
            article.year,
            article.title_article,
            article.publisher,
           	article.type,
            article.risc,
            article.issn,
            article.edn,
            authors_organisations.author_id,
			authors_organisations.author_name,
			authors_organisations.author_initials,
            authors_organisations.org_id,
            authors_organisations.org_name,
			cte.author_count,
			COUNT(author_id) over affilations_cnt as affilations_count
        FROM
            article 
		JOIN
			article_authors_linkage USING(item_id)
		JOIN 
			authors_organisations USING(counter)
		JOIN
			cte USING(linkurl)
		window affilations_cnt as (partition by linkurl, author_id)) AS subquery
		WHERE {0}
            """.format(where[0])

                cursor.execute(sql_query)
                result = cursor.fetchall()
                columns = [desc[0] for desc in cursor.description]
                df = pd.DataFrame(result, columns=columns)
                cursor.close()
                connection.close()

                excel_template_path = "shablon_kbpr.xlsx"
                df_template = pd.read_excel(excel_template_path)
                df_template['URL'] = df['linkurl']
                df_template['Идентификатор DOI *'] = df['doi']
                df_template['Количество авторов *'] = df['author_count']
                df_template['Фамилия *'] = df['author_name']
                df_template['Имя *'] = df['author_initials']
                # df_template['Отчество'] = df['patronymic']
                # df_template['Должность *'] = df['position']
                # df_template['Ученая степень *'] = df['academic_degree']
                # df_template['Тип трудовых отношений *'] = df['employment_relationship']
                # df_template['Год рождения *'] = df['birth_year']
                df_template['Количество аффиляций *'] = df['affilations_count']
                df_template['Аффиляция *'] = df['org_name']
                df_template['Дата публикации *'] = pd.to_datetime(df['year'], format='%Y').dt.strftime('01/01/%Y')
                df_template['Наименование публикации *'] = df['title_article']
                df_template['Наименование издания *'] = df['publisher']
                df_template['Вид издания  *'] = df['type']
                df_template['Идентификатор РИНЦ'] = df['risc']
                df_template['Идентификатор ISSN'] = df['issn']
                df_template['Идентификатор EDN'] = df['edn']

                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                output_path = f"shablon_kbpr_{timestamp}.xlsx"
                df_template.to_excel(output_path)
                QMessageBox.information(self, "Экспорт", "Excel файл по шаблону кбпр создан!")

            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
            finally:
                try:
                    cursor.close()
                except:
                    pass
                try:
                    connection.close()
                except:
                    pass


    def execute_query_with_params(self,query):
        query = query
        conn = psycopg2.connect(
            dbname=database_parametres['dbname'],
            user=database_parametres['user'],
            password=database_parametres['password'],
            host=database_parametres['host'],
            port=database_parametres['port']
        )
        cur = conn.cursor()

        cur.execute(query)

        conn.commit()

        cur.close()
        conn.close()



    splitInitials_query = """
        DROP TABLE IF EXISTS authors_splitted;
        CREATE TABLE  authors_splitted AS
        SELECT DISTINCT item_id,author_id, author_name,author_initials,
            CASE
                WHEN author_initials LIKE '% %' THEN split_part(whole_table.author_initials, ' ', 1)
                WHEN author_initials NOT LIKE '% %' AND author_initials NOT LIKE '%.%' THEN author_initials
                WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 2 THEN author_initials
                WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 4 THEN LEFT(author_initials, 2)
				WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 5 THEN LEFT(author_initials, 2)
                ELSE author_initials
            END AS first_name,
            CASE
                WHEN author_initials LIKE '% %' THEN split_part(whole_table.author_initials, ' ', -1)
                WHEN author_initials NOT LIKE '% %' AND author_initials NOT LIKE '%.%' THEN NULL
                WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 2 THEN NULL
                WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 4 THEN RIGHT(author_initials, 2)
				WHEN author_initials LIKE '%.%' AND LENGTH(author_initials) = 5 THEN RIGHT(author_initials, 3)
                ELSE author_initials
            END AS patronymic
        FROM whole_table;
        """
    createAuthorsReference_query = """
        DROP TABLE IF EXISTS authors_reference_with_id;
        CREATE TABLE authors_reference_with_id AS
        SELECT author_id,ar."Автор публикации" AS publication_author,at.lastname,at.first_name,at.patronymic,
        ar."Должность автора статьи в организ" AS position,
        ar."Ученая степень" AS academic_degree ,ar."Тип трудовых отношений" AS employment_relationship,ar."Год рождения автора" AS birth_year
        FROM authors_splitted AS at
	    INNER JOIN authors_reference AS ar
	    ON (at.lastname,at.first_name,at.patronymic) = (ar."Фамилия",ar."имя",ar."отчество")
        WHERE author_id IS NOT NULL
        UNION
        SELECT at.author_id,at.full_name,ar."Фамилия",ar."имя",ar."отчество",ar."Должность автора статьи в организ",
        ar."Ученая степень",ar."Тип трудовых отношений",ar."Год рождения автора"
        FROM (SELECT lastname || ' ' || initials as full_name,author_id FROM authors
        WHERE initials LIKE '%.%' AND LENGTH(initials)  = 4 AND author_id IS NOT NULL) AS at
	    INNER JOIN authors_reference AS ar
	    ON (at.full_name) = (ar."Автор публикации")
        WHERE author_id IS NOT NULL
        """


    def import_xlsx_to_postgresql2(self, database_params, xlsx_file_path, table_name, index_col):
        try:
            data_from_sql = []
            data_from_excel = []
            index_sql = []
            index_excel = []
            connection_str = f"postgresql://{database_params['user']}:{database_params['password']}@{database_params['host']}:{database_params['port']}/{database_params['dbname']}"
            engine = create_engine(connection_str)
            def replace_float_with_null(value):
                if isinstance(value, float):
                    return None
                return value
            float_columns = [
                'linkurl',
                'genre',
                'type',
                'journal_title',
                'issn',
                'eissn',
                'publisher',
                'vak',
                'wos',
                'scopus',
                'number',
                'page_begin',
                'page_end',
                'language',
                'title_article',
                'doi',
                'edn',
                'risc',
                'corerisc']
            data_frame = pd.read_excel(xlsx_file_path, index_col=index_col)
            if table_name == 'authors_organisations':
                if "Unnamed: 0" in data_frame.columns:
                    data_frame = data_frame.drop("Unnamed: 0", axis=1)
                data_frame.drop("author_fullname", axis=1, inplace=True)
            if table_name == 'article':
                for column in float_columns:
                    data_frame[column] = data_frame[column].apply(lambda x: replace_float_with_null(x))
            existing_data_query = f"SELECT * FROM {table_name}"
            existing_data = pd.read_sql(existing_data_query, engine)
            if table_name == 'article':
                columns_to_compare = ['item_id', 'linkurl', 'genre', 'type', 'issn', 'eissn', 'publisher', 'vak', 'rcsi', 'wos', 'scopus', 'quartile', 'year', 'number', 'contnumber', 'volume', 'language', 'edn', 'grnti', 'risc', 'corerisc', 'doi']
                new_row = pd.DataFrame({'item_id': ''}, index=[0])
                data_frame = pd.concat([data_frame, new_row])
                merged_data = pd.concat([data_frame, existing_data])
                merged_data = merged_data.drop_duplicates()
                merged_data.to_excel('merged.xlsx')
                update_excel_file('merged.xlsx')
                merged_data = pd.read_excel('merged.xlsx', index_col=0)
                duplicate_rows = merged_data.duplicated(subset=columns_to_compare, keep=False)
                duplicate_data = merged_data[duplicate_rows]
                # duplicate_data = duplicate_data.sort_values(by=['item_id'])
                duplicate_data.to_excel('duplicate.xlsx')
                for index, row in duplicate_data.iterrows():
                    if row['data_origin'] == 'sql':
                        index_sql.append(index)
                        data_from_sql.append(row[['item_id','linkurl', 'genre', 'type', 'journal_title', 'issn', 'eissn', 'publisher', 'vak',	'rcsi', 'wos', 'scopus', 'quartile', 'year', 'number', 'contnumber', 'volume', 'page_begin', 'page_end', 'language',
                                         'doi',	'edn', 'grnti', 'risc', 'corerisc']].values)
                    elif row['data_origin'] == 'excel':
                        index_excel.append(index)
                        data_from_excel.append(row[['item_id', 'linkurl', 'genre', 'type', 'journal_title', 'issn', 'eissn',
                                        'publisher', 'vak', 'rcsi', 'wos', 'scopus', 'quartile', 'year', 'number',
                                        'contnumber', 'volume', 'page_begin', 'page_end', 'language',
                                        'doi', 'edn', 'grnti', 'risc', 'corerisc']].values)
                if(len(data_from_sql) > 0):
                    self.showDialog(data_from_sql, data_from_excel, index_sql, index_excel)
                    merged_data = merged_data[~((merged_data.index.isin(index_sql)) & (merged_data['data_origin'] == 'sql'))]
                    merged_data = merged_data[~((merged_data.index.isin(index_excel)) & (merged_data['data_origin'] == 'excel'))]
                merged_data.drop("data_origin", axis=1, inplace=True)
                merged_data = merged_data.dropna(how='all', subset=['item_id', 'linkurl', 'genre', 'type', 'journal_title', 'issn', 'eissn',
                                        'publisher', 'vak', 'rcsi', 'wos', 'scopus', 'quartile', 'year', 'number',
                                        'contnumber', 'volume', 'page_begin', 'page_end', 'language',
                                        'doi', 'edn', 'grnti', 'risc', 'corerisc'])
                if "Unnamed: 0" in merged_data.columns:
                    merged_data = merged_data.drop("Unnamed: 0", axis=1)
                merged_data.to_excel('merged.xlsx')
                merged_data = pd.read_excel('merged.xlsx', index_col=0)
                merged_data.to_sql(table_name, engine, if_exists='replace', index=False)
                fix_query = f"SELECT * FROM {table_name}"
                data_test = pd.read_sql(fix_query, engine)
                data_test = data_test.drop_duplicates()
                data_test.to_sql(table_name, engine, if_exists='replace', index=False)
            elif table_name == 'authors_organisations':
                merged_data = pd.concat([data_frame, existing_data])
                merged_data = merged_data.drop_duplicates()
                merged_data.to_excel('merged_ao.xlsx')
                deduplicate_excel('merged_ao.xlsx')
                merged_data_filtered = pd.read_excel('merged_ao.xlsx')
                merged_data_filtered.to_sql(table_name, engine, if_exists='replace', index=False)
            elif table_name == 'article_authors_linkage':
                merged_data = pd.concat([data_frame, existing_data])
                merged_data = merged_data.drop_duplicates()
                merged_data.to_excel('merged_link.xlsx')
                merged_data.to_sql(table_name, engine, if_exists='replace', index=False)
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            pass
    def importButtonClickHandler(self):
        self.ui.progressBar.setValue(0)
        fname = QFileDialog.getOpenFileName(self, "Open XML file", "", "All Files (*);; XML Files (*.xml)")
        if fname[0]:
            parse_articles_to_excel(fname[0])
            self.ui.progressBar.setValue(10)
            data = update_elibrary_id('authors_organisations_initial.xlsx')
            self.showDialog_2(data)
            self.ui.progressBar.setValue(20)
            self.ui.progressBar.setValue(30)
            update_org_id('authors_organisations.xlsx')
            update_author_id('authors_organisations.xlsx')
            deduplicate_excel('authors_organisations.xlsx')
            self.ui.progressBar.setValue(40)
            self.ui.progressBar.setValue(50)
            self.import_xlsx_to_postgresql2(database_parametres, 'authors_organisations.xlsx', 'authors_organisations', False)
            self.ui.progressBar.setValue(60)
            self.ui.progressBar.setValue(70)
            self.ui.progressBar.setValue(80)
            self.ui.progressBar.setValue(90)
            self.import_xlsx_to_postgresql2(database_parametres, 'article.xlsx', 'article', None)
            self.import_xlsx_to_postgresql2(database_parametres, 'article_authors_linkage.xlsx', 'article_authors_linkage', None)
            self.ui.progressBar.setValue(100)
            QMessageBox.information(self, "Успешный импорт", "Данные были перенесены в Базу Данных!")
        else:
            print("Выбор файла отменен. Файл не был перемещен.")

    def searchButtonDBConnector(self, year, lastname):
        query = """
                        SELECT item_id, author_name, linkurl, genre, type, journal_title,publisher, title_article
                        FROM article
                        JOIN authors_organisations USING(counter)
                        WHERE year = '{year}' AND author_name = '{lastname}'
                        """

        query = query.format(year=year, lastname=lastname)
        conn = psycopg2.connect(database=database_parametres['dbname'],
                                user=database_parametres['user'],
                                password=database_parametres['password'],
                                host=database_parametres['host'],
                                port=database_parametres['port'])
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()
        self.ui.tableWidget.setRowCount(0)

        for row_number, row_data in enumerate(result):
            self.ui.tableWidget.insertRow(row_number)
            for column_number, data in enumerate(row_data):
             self.ui.tableWidget.setItem(row_number, column_number, QtWidgets.QTableWidgetItem(str(data)))
        cur.close()
        conn.close()

    def userChoicePatternFetchFromDB(self,columns):
        query = """
        SELECT DISTINCT {columns}
        FROM authors_splitted
        JOIN authors_organisations ON CAST(authors_splitted.author_id AS text) = authors_organisations.author_id
                                   OR (authors_splitted.author_id IS NULL AND authors_organisations.author_id IS NULL)
                                   AND authors_splitted.lastname = authors_organisations.author_name
        JOIN article_author ON CAST(article_author.author_id AS text) = authors_organisations.author_id
                            OR (article_author.author_id IS NULL AND authors_organisations.author_id IS NULL)
                            AND authors_organisations.author_name = article_author.author_name
        JOIN article ON article.item_id = article_author.item_id
        LEFT JOIN authors_reference_with_id ON CAST(authors_reference_with_id.author_id AS text) = authors_organisations.author_id
        JOIN (
            SELECT item_id, COUNT(author_name) AS author_count
            FROM article_author
            GROUP BY item_id
        ) AS nested_auth ON article.item_id = nested_auth.item_id
        JOIN (
            SELECT article.item_id, COUNT(author_name) AS aff_count
            FROM article
            INNER JOIN article_author ON article.item_id = article_author.item_id
            GROUP BY doi, article.item_id
        ) AS nested_aff ON article.item_id = nested_aff.item_id
        WHERE authors_organisations.org_id = '570'
                           """
        query = query.format(columns=columns)
        conn = psycopg2.connect(database=database_parametres['dbname'],
                                user=database_parametres['user'],
                                password=database_parametres['password'],
                                host=database_parametres['host'],
                                port=database_parametres['port'])
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()
        columns = columns.split(",")
        df = pd.DataFrame(result, columns=columns)
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_path = f"userTemplate_{timestamp}.xlsx"
        df.to_excel(output_path, index=False, sheet_name='Sheet1')
        QMessageBox.information(self, "Экспорт", "Excel файл по шаблону пользователя создан!")

    def exportConnector(self, year, lastname):
        query = """
                        SELECT a.item_id, aa.author_name, a.linkurl, a.genre, a.type, a.journal_title, a.publisher, a.title_article
                        FROM article AS a
                        INNER JOIN article_author AS aa ON aa.item_id = a.item_id
                        WHERE a.year = '{year}' AND aa.author_name = '{lastname}'
                        """

        query = query.format(year=year, lastname=lastname)
        conn = psycopg2.connect(database=database_parametres['dbname'],
                                user=database_parametres['user'],
                                password=database_parametres['password'],
                                host=database_parametres['host'],
                                port=database_parametres['port'])
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()

    def get_test_auf(self):
        text_array = []

        for combobox in self.ui.comboboxes:
            current_text = combobox.currentText()
            if current_text != 'None':
                if current_text in ["item_id", "linkurl", "genre", "type", "journal_title", "issn", "eissn",
                                    "publisher", "vak", "rcsi", "wos", "scopus", "quartile", "year", "number",
                                    "contnumber", "volume", "page_begin", "page_end", "language",
                                    "title_article", "doi", "edn", "grnti", "risc", "corerisc"]:
                    text_array.append("article." + current_text)
                elif current_text == "last_name":
                    text_array.append(
                        "CASE WHEN authors_splitted.lastname ~ '[A-Za-z]' AND authors_reference_with_id.birth_year IS NOT NULL THEN authors_reference_with_id.lastname ELSE authors_splitted.lastname END AS last_name")
                elif current_text == "first_name":
                    text_array.append(
                        "CASE WHEN (authors_splitted.first_name LIKE '%.%' AND authors_reference_with_id.birth_year IS NOT NULL) OR (authors_splitted.first_name ~ '[A-Za-z]' AND authors_reference_with_id.birth_year IS NOT NULL) OR authors_splitted.first_name IS NULL THEN authors_reference_with_id.first_name ELSE authors_splitted.first_name END AS first_name")
                elif current_text == "patronymic":
                    text_array.append(
                        "CASE WHEN (authors_splitted.patronymic LIKE '%.%' AND authors_reference_with_id.birth_year IS NOT NULL) OR authors_splitted.patronymic IS NULL OR (authors_splitted.patronymic ~ '[A-Za-z]'  AND authors_reference_with_id.birth_year IS NOT NULL) THEN authors_reference_with_id.patronymic ELSE authors_splitted.patronymic END AS patronymic")
                elif current_text in ["position", "degree", "employment_relationship",
                        "birth_year"]:
                    text_array.append("authors_reference_with_id." + current_text)
                elif current_text == "author_count":
                    text_array.append("nested_auth." + current_text)
                elif current_text == "aff_count":
                    text_array.append("nested_aff." + current_text)
                elif current_text == "org_id":
                    text_array.append("authors_organisations." + current_text)
                elif current_text == "org_name":
                    text_array.append("authors_organisations." + current_text)
                else:
                    text_array.append(current_text)
        result = ','.join(text_array)
        result = result.split(",")
        result = pd.Series(result).drop_duplicates().tolist()
        result = ','.join(result)
        self.userChoicePatternFetchFromDB(result)

    def get_text(self):
        text = self.ui.textEdit.toPlainText().strip()
        selected_text = self.ui.comboBox.currentText()
        self.searchButtonDBConnector(selected_text, text)

    def getYearAndSurname(self):
        specify_where_basic = []
        specify_where_advanced = []
        text = self.ui.textEdit_2.text()
        selected_year_from = self.ui.comboBox_2.currentText()
        selected_year_to = self.ui.comboBox_3.currentText()
        if selected_year_from > selected_year_to:
            QMessageBox.information(self, "Ошибка", "Проверьте диапозон годов")
            return
        else:
            if selected_year_from == 'None' and selected_year_to == 'None' and text == '':
             pass
            elif (selected_year_from != 'None' and selected_year_to == 'None' and text == ''):
             specify_where_basic.append(f" subquery.year = {selected_year_from}")
            elif (selected_year_from != 'None' and selected_year_to != 'None' and text == ''):
             specify_where_basic.append(f" year BETWEEN {selected_year_from}  AND  {selected_year_to} ")
            elif selected_year_from == 'None' and selected_year_to == 'None' and text != '':
                specify_where_advanced.append(f" author_id IN (SELECT  author_id FROM authors_organisations WHERE  author_name || ' ' || author_initials = '{text} ')")
            elif (selected_year_from != 'None' and selected_year_to != 'None' and text != ''):
             specify_where_advanced.append(f" year BETWEEN {selected_year_from} AND {selected_year_to} AND author_id IN (SELECT author_id FROM authors_organisations WHERE author_name || ' ' || author_initials = '{text} ')")
            elif (selected_year_from != 'None' and selected_year_to == 'None' and text != ''):
             specify_where_advanced.append(f" year = {selected_year_from} AND author_id IN (SELECT  author_id FROM authors_organisations WHERE  author_name || ' ' || author_initials = '{text} ')")
        if len(specify_where_basic) > 0 or (len(specify_where_basic) == 0 and len(specify_where_advanced) == 0 ):
            print(specify_where_basic)
            self.process_data(specify_where_basic)
        else:
            self.process_data_advanced(specify_where_advanced)
            print(specify_where_advanced)

    def addOneRowToDB(self):
        for row in range(self.ui.tableWidget_add_row.rowCount()):
            row_data = []
            for column in range(self.ui.tableWidget_add_row.columnCount()):
                item = self.ui.tableWidget_add_row.item(row, column)
                print(item)
                if item is not None:
                    cell_data = item.text()
                    row_data.append(cell_data)
                else:
                    row_data.append("NULL")
            print(row_data)
            self.insertNewRowInWholeTable(row_data)
            QMessageBox.information(self, "Успешно", "Строка была добавлена в базу данных!")

    def insertNewRowInWholeTable(self, row_1):
        try:
            query = """
                INSERT INTO article (item_id, linkurl, genre, type, journal_title, issn, eissn, publisher, vak, rcsi, wos, scopus, quartile,
                year, number, contnumber, volume, page_begin, page_end, language, title_article, doi, edn, grnti, risc, corerisc, counter) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
            """
            print(len(row_1))
            conn = psycopg2.connect(database=database_parametres['dbname'],
                                    user=database_parametres['user'],
                                    password=database_parametres['password'],
                                    host=database_parametres['host'],
                                    port=database_parametres['port'])
            cur = conn.cursor()
            cur.execute(query, row_1)
            conn.commit()

        except psycopg2.Error as e:
            print(f"Error: {e}")

        except Exception as e:
            print(f"Unexpected Error: {e}")

        finally:
            if cur:
                cur.close()
            if conn:
                conn.close()

    def authorsReferenceToSQL(self,database_params):
        fname = QFileDialog.getOpenFileName(self, "Open XML file", "", "All Files (*);; XML Files (*.xml)")
        if fname[0]:
            connection_str = f"postgresql://{database_params['user']}:{database_params['password']}@{database_params['host']}:{database_params['port']}/{database_params['dbname']}"
            engine = create_engine(connection_str)
            data_frame = pd.read_excel(fname[0])
            data_frame.to_sql('authors_reference', engine, index=False, if_exists='replace')
            QMessageBox.information(self, "Успешный импорт", "Данные были перенесены в Базу Данных!")
        else:
            print("Выбор файла отменен. Файл не был перемещен.")

    def on_home_button_iconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(0)

    def on_home_button_iconexpandedwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(0)

    def on_articleDB_button_iconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(1)


    def on_articleDB_button_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(1)

    def on_article_authorDB_button_iconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(2)

    def on_addingdatatoBD_button_iconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(2)

    def on_addingdatatoBD_button_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(2)

    def on_export_button_onlyiconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(4)

    def on_export_button_expandedwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(4)

    def on_import_button_onlyiconwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(3)

    def on_import_button_expandedwidget_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(3)

    def on_pushButton_2_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(5)

    def on_pushButton_5_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(5)

    def on_article_authorDB_button_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(6)

    def on_authorsDB_button_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(7)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    with open("style.qss", "r") as style_file:
        style_str = style_file.read()
    app.setStyleSheet(style_str)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())