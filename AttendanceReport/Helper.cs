using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;


namespace AttendanceReport
{
    public static class Helper
    {
        public const string CONST_ENC_PASSPHRASE = "Pas5pr@se";
        public const string CONST_ENC_SALT_VALUE = "s@1tValue";
        public const string CONST_ENC_HASH_ALGO = "MD5";
        public const int CONST_ENC_PASSWORD_ITERATION = 3;
        public const string CONST_ENC_INIT_VECTOR = "@1B2c3D4e5F6g7H8";
        public const int CONST_ENC_KEY_SIZE = 256;

        public static string EncryptString(string plainText,
           string passPhrase,
           string saltValue,
           string hashAlgorithm,
           int passwordIterations,
           string initVector,
           int keySize)
        {
            // Convert strings into byte arrays.
            // Let us assume that strings only contain ASCII codes.
            // If strings include Unicode characters, use Unicode, UTF7, or UTF8 
            // encoding.
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

            // Convert our plaintext into a byte array.
            // Let us assume that plaintext contains UTF8-encoded characters.
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            // First, we must create a password, from which the key will be derived.
            // This password will be generated from the specified passphrase and 
            // salt value. The password will be created using the specified hash 
            // algorithm. Password creation can be done in several iterations.
            PasswordDeriveBytes password = new PasswordDeriveBytes(
                passPhrase,
                saltValueBytes,
                hashAlgorithm,
                passwordIterations);

            // Use the password to generate pseudo-random bytes for the encryption
            // key. Specify the size of the key in bytes (instead of bits).
            byte[] keyBytes = password.GetBytes(keySize / 8);

            // Create uninitialized Rijndael encryption object.
            RijndaelManaged symmetricKey = new RijndaelManaged();

            // It is reasonable to set encryption mode to Cipher Block Chaining
            // (CBC). Use default options for other symmetric key parameters.
            symmetricKey.Mode = CipherMode.CBC;

            // Generate encryptor from the existing key bytes and initialization 
            // vector. Key size will be defined based on the number of the key 
            // bytes.
            ICryptoTransform encryptor = symmetricKey.CreateEncryptor(
                keyBytes,
                initVectorBytes);

            // Define memory stream which will be used to hold encrypted data.
            MemoryStream memoryStream = new MemoryStream();
            byte[] cipherTextBytes;

            // Define cryptographic stream (always use Write mode for encryption).
            using (CryptoStream cryptoStream = new CryptoStream(memoryStream,
                      encryptor,
                      CryptoStreamMode.Write))
            {
                // Start encrypting.
                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);

                // Finish encrypting.
                cryptoStream.FlushFinalBlock();

                // Convert our encrypted data from a memory stream into a byte array.
                cipherTextBytes = memoryStream.ToArray();

                // Close both streams.
                memoryStream.Close();
                cryptoStream.Close();
            }

            // Convert encrypted data into a base64-encoded string.
            string cipherText = Convert.ToBase64String(cipherTextBytes);

            // Return encrypted string.
            return cipherText;
        }


        /// <summary>
        /// Decrypts the string.
        /// </summary>
        /// <param name="cipherText">The cipher text.</param>
        /// <param name="passPhrase">The pass phrase.</param>
        /// <param name="saltValue">The salt value.</param>
        /// <param name="hashAlgorithm">The hash algorithm.</param>
        /// <param name="passwordIterations">The password iterations.</param>
        /// <param name="initVector">The init vector.</param>
        /// <param name="keySize">Size of the key.</param>
        /// <returns></returns>
        public static string DecryptString(string cipherText,
            string passPhrase,
            string saltValue,
            string hashAlgorithm,
            int passwordIterations,
            string initVector,
            int keySize)
        {
            // Mustafa: If there is null or empty text then return empty string.
            if (string.IsNullOrEmpty(cipherText))
            {
                return string.Empty;
            }

            // Convert strings defining encryption key characteristics into byte
            // arrays. Let us assume that strings only contain ASCII codes.
            // If strings include Unicode characters, use Unicode, UTF7, or UTF8
            // encoding.
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

            // Convert our ciphertext into a byte array.
            byte[] cipherTextBytes = Convert.FromBase64String(cipherText);

            // First, we must create a password, from which the key will be 
            // derived. This password will be generated from the specified 
            // passphrase and salt value. The password will be created using
            // the specified hash algorithm. Password creation can be done in
            // several iterations.
            PasswordDeriveBytes password = new PasswordDeriveBytes(
                passPhrase,
                saltValueBytes,
                hashAlgorithm,
                passwordIterations);

            // Use the password to generate pseudo-random bytes for the encryption
            // key. Specify the size of the key in bytes (instead of bits).
            byte[] keyBytes = password.GetBytes(keySize / 8);

            // Create uninitialized Rijndael encryption object.
            RijndaelManaged symmetricKey = new RijndaelManaged();

            // It is reasonable to set encryption mode to Cipher Block Chaining
            // (CBC). Use default options for other symmetric key parameters.
            symmetricKey.Mode = CipherMode.CBC;

            // Generate decryptor from the existing key bytes and initialization 
            // vector. Key size will be defined based on the number of the key 
            // bytes.
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(
                keyBytes,
                initVectorBytes);

            // Define memory stream which will be used to hold encrypted data.
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
            string plainText = string.Empty;

            // Define cryptographic stream (always use Read mode for encryption).
            using (CryptoStream cryptoStream = new CryptoStream(memoryStream,
                      decryptor,
                      CryptoStreamMode.Read))
            {
                // Since at this point we don't know what the size of decrypted data
                // will be, allocate the buffer long enough to hold ciphertext;
                // plaintext is never longer than ciphertext.
                byte[] plainTextBytes = new byte[cipherTextBytes.Length];

                // Start decrypting.
                int decryptedByteCount = cryptoStream.Read(plainTextBytes,
                    0,
                    plainTextBytes.Length);

                // Close both streams.
                memoryStream.Close();
                cryptoStream.Close();


                // Convert decrypted data into a string. 
                // Let us assume that the original plaintext string was UTF8-encoded.
                plainText = Encoding.UTF8.GetString(plainTextBytes,
                    0,
                    decryptedByteCount);
            }

            // Return decrypted string.   
            return plainText;
        }

    }
}
