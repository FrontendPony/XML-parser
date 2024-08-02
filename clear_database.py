import psycopg2
from dbsettings import database_parametres as database_parameters

def delete_data():
    query = (
        "DELETE FROM article;"
        "DELETE FROM authors_organisations;"
        "DELETE FROM article_authors_linkage;"
        "DELETE FROM alternative_author_ids;"
        "DELETE FROM conference;"
    )

    try:
        conn = psycopg2.connect(
            dbname=database_parameters['dbname'],
            user=database_parameters['user'],
            password=database_parameters['password'],
            host=database_parameters['host'],
        )

        cur = conn.cursor()

        cur.execute(query)

        conn.commit()

    except (Exception, psycopg2.DatabaseError) as error:
        print("Error deleting data:", error)

    finally:
        cur.close()
        conn.close()

delete_data()



