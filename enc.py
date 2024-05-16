from cryptography.fernet import Fernet

# Generate a key
key = Fernet.generate_key()

# Save the key to a file
with open('encryption_key.key', 'wb') as key_file:
    key_file.write(key)

# Create a Fernet symmetric key
cipher = Fernet(key)

# Read the contents of the clean output CSV file
input_file_path = r"C:\Users\ASUS\Desktop\Security__project\monoprix_boissons.csv"
with open(input_file_path, 'rb') as file:
    data = file.read()

# Encrypt the data
encrypted_data = cipher.encrypt(data)

# Write the encrypted data to a new file
output_file_path = r"C:\Users\ASUS\Desktop\Security__project\monoprix_boissons.csv"
with open(output_file_path, 'wb') as file:
    file.write(encrypted_data)

print("File encrypted successfully.")
