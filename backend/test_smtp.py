import smtplib

smtp_server = "smtp.gmail.com"
port = 587
username = "lca.platformlaura@gmail.com"
password = "ufafnedoucnnrhqq"   # copiez-collez le mot de passe exact

try:
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(username, password)
    print("✅ Connexion SMTP réussie")
    server.quit()
except Exception as e:
    print("❌ Erreur de connexion :", e)