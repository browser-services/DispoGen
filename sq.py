import sqlite3

DB_PATH = "users_reports.db"

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# Create Users table
cur.execute("""
CREATE TABLE IF NOT EXISTS Users (
    id INTEGER PRIMARY KEY ,
    full_name TEXT NOT NULL
)
""")

# Create Reports table
cur.execute("""
CREATE TABLE IF NOT EXISTS Reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    text TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES Users(id)
)
""")

# cur.execute("""
#             INSERT INTO Users (`f`ull_name) VALUES
#             ("Kerwin Ray Abalos"),
#             ("Ethan Hunt"),
#             ("John Wick"),
#             ("Uchiha Sasuke")
#             """)

cur.execute("""
            INSERT INTO Users (full_name) VALUES
            ("John Paul Santos"),
            ("Nelson Anzano Oligario"),
            ("Kerwin Ray Abalos"),
            ("Camille Gutierrez Gong"),
            ("Ethan Aguinaldo"),
            ("Ernest Bustamante Aquino"),
            ("Mico Mendoza")
            """)

# cur.execute("""
#             INSERT INTO Reports (id, user_id, text) VALUES
#             (1, 1, "Studied Python"),
#             (2, 1, "Reviewing and practicing TypeScript"),
#             (3, 1, "Continued redesigning the UI/UX of the Daily News Scraper"),
#             (4, 1, "Continued studying Next.js")
#             """)

conn.commit()
conn.close()
print("Database initialized.")
