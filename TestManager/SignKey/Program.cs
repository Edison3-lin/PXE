using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace SignKey {
    public class DigitalSignature {
        public RSAParameters PublicKey { get; private set; }
        public RSAParameters PrivateKey { get; private set; }

        public DigitalSignature()
        {
            GenerateKeys();
        }

        private void GenerateKeys()
        {
            using (var provider = new RSACryptoServiceProvider(2048))
            {
                PrivateKey = provider.ExportParameters(true);
                PublicKey = provider.ExportParameters(false);
            }
        }

        public byte[] SignData(string data)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportParameters(PrivateKey);
                var dataBytes = Encoding.UTF8.GetBytes(data);
                return rsa.SignData(dataBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            }
        }

        public bool VerifySignature(string data, byte[] signature)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportParameters(PublicKey);
                var dataBytes = Encoding.UTF8.GetBytes(data);
                return rsa.VerifyData(dataBytes, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            }
        }

        public void SaveKeyToFile(string fileName, RSAParameters key, bool includePrivateParameters)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportParameters(key);
                string keyString = rsa.ToXmlString(includePrivateParameters);
                File.WriteAllText(fileName, keyString);
            }
        }

        public RSAParameters LoadKeyFromFile(string fileName, bool includePrivateParameters)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                string keyString = File.ReadAllText(fileName);
                rsa.FromXmlString(keyString);
                return rsa.ExportParameters(includePrivateParameters);
            }
        }
        // ============== MAIN ==============
        static void Main(string[] args) {

            // 使用範例
            var digitalSignature = new DigitalSignature();

            // 簽名
            // string fileContent = File.ReadAllText("path/to/your/file.txt");
            // byte[] signature = digitalSignature.SignData(fileContent);

            // // 驗證簽名
            // bool isVerified = digitalSignature.VerifySignature(fileContent, signature);
            // Console.WriteLine("Signature Verified: " + isVerified);

            // 保存密鑰
            digitalSignature.SaveKeyToFile("privateKey.xml", digitalSignature.PrivateKey, true);
            digitalSignature.SaveKeyToFile("publicKey.xml", digitalSignature.PublicKey, false);

            // // 加載密鑰
            // var loadedPrivateKey = digitalSignature.LoadKeyFromFile("privateKey.xml", true);
            // var loadedPublicKey = digitalSignature.LoadKeyFromFile("publicKey.xml", false);
        }
    }
}