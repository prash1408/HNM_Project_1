import sqlite3
conn = sqlite3.connect('Requirement.db')
c= conn.cursor()
c.execute(f'''
    CREATE TABLE IF NOT EXISTS Nomura_NonTECH (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        req_id TEXT,
        req_Name TEXT,
        date TEXT,
        sal_budget TEXT,
        Notice_Period TEXT,
        Mandatory_skills TEXT,
        Desirable_Skills TEXT,
        exp_range TEXT,
        Client_Spoc TEXT,
        Total_Upload TEXT,
        Remarks TEXT ,
        Location TEXT,
        Status TEXT 
    )
''')    
conn.commit()
conn.close()



# 'Nomura_JAVA', 'British_Petrolium', 'Nomura_TECH', 'Morgan_Stanley', 'Russell', 'MUFG','Chevron','Lufthansa','Interactive_Brokers'

# c.execute("ALTER TABLE Nomura_JAVA ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE British_Petrolium ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Nomura_TECH ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Morgan_Stanley ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Russell ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE MUFG ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Chevron ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Lufthansa ADD COLUMN Status TEXT")
# c.execute("ALTER TABLE Interactive_Brokers ADD COLUMN Status TEXT")