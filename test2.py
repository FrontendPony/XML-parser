import psycopg2
from dbsettings import database_parametres as database_parameters

query = f"DELETE FROM article;DELETE  FROM authors_organisations;DELETE FROM article_authors_linkage;DELETE FROM alternative_author_ids;"


# Connect to the database
conn = psycopg2.connect(
    dbname=database_parameters['dbname'],
    user=database_parameters['user'],
    password=database_parameters['password'],
    host=database_parameters['host'],
)

cur = conn.cursor()

# Execute the DELETE query
cur.execute(query)

# Commit the changes to the database
conn.commit()

# Close the cursor and database connection
cur.close()
conn.close()


