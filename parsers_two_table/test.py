import pandas as pd
from sqlalchemy import create_engine
from dbsettings import database_parametres as database_params

connection_str = f"postgresql://{database_params['user']}:{database_params['password']}@{database_params['host']}:{database_params['port']}/{database_params['dbname']}"
engine = create_engine(connection_str)
data_frame = pd.read_excel('article.xlsx', index_col=0)
# data_frame.drop("author_fullname", axis=1, inplace=True)
data_frame.to_sql('article', engine, if_exists='replace')


