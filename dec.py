from cryptography.fernet import Fernet

# Load the key (replace 'key.txt' with the path to your key file)
with open('encryption_key.key', 'rb') as key_file:
    key = key_file.read()

# Create a Fernet symmetric key
cipher = Fernet(key)

# Read the contents of the encrypted file
input_file_path = r"C:\Users\ASUS\Desktop\Security__project\monoprix_boissons.csv"
with open(input_file_path, 'rb') as file:
    encrypted_data = file.read()

# Decrypt the data
decrypted_data = cipher.decrypt(encrypted_data)

# Write the decrypted data to a new file
output_file_path = r"C:\Users\ASUS\Desktop\Security__project\monoprix_boissons_dec.csv"
with open(output_file_path, 'wb') as file:
    file.write(decrypted_data)

print("File decrypted successfully.")