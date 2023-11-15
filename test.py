import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine
from dbsettings import database_parametres as database_params

connection_str = f"postgresql://{database_params['user']}:{database_params['password']}@{database_params['host']}:{database_params['port']}/{database_params['dbname']}"
engine = create_engine(connection_str)
data_frame = pd.read_excel('article_authors_linkage.xlsx')
# data_frame.drop("author_fullname", axis=1, inplace=True)
print(data_frame)
data_frame.to_sql('article_authors_linkage', engine, if_exists='replace', index=False)


