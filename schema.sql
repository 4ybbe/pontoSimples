
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome BLOB NOT NULL,
            cpf BLOB UNIQUE NOT NULL,
            otp_secret BLOB NOT NULL
        );
        