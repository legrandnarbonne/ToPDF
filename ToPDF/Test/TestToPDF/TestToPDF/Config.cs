using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;

namespace ToPDF.Classes
{
    [Serializable]
    public class Config : ICloneable
    {

        private string passphrase = "emfklvkzdc12";
        private string passwordStore;

        #region properties

        public WorkingConfig[] WorkingConfigList { get; set; }

        #endregion properties


        /// <summary>
        /// Return application path
        /// </summary>
        public static string DefaultPath
        {
            get { return AppDomain.CurrentDomain.BaseDirectory; }
        }


        /// <summary>
        /// Return default config file path
        /// </summary>
        public static string DefaultFileName
        {
            get { return DefaultPath + "\\default.conf"; }
        }
        

        [OnSerializing()]
        private void OnSerializingMethod(StreamingContext context)
        {
            //passwordStore = Password;
            //if (!string.IsNullOrEmpty(Password)) Password = Encrypt.EncryptString(Password, passphrase);
        }

        [OnSerialized()]
        private void OnSerializedMethod(StreamingContext context)
        {
            //Password = passwordStore;
        }

        [OnDeserialized()]
        private void OnDeserializedMethod(StreamingContext context)
        {
            //if (!string.IsNullOrEmpty(Password)) Password = Encrypt.DecryptString(Password, passphrase);
        }

        #region method

        /// <summary>
        /// Saves to an xml file
        /// </summary>
        /// <param name="FileName">File path of the new xml file</param>
        public void Save(string FileName)
        {
            using (var writer = new System.IO.StreamWriter(FileName))
            {
                var serializer = new XmlSerializer(this.GetType());
                serializer.Serialize(writer, this);
                writer.Flush();
            }
        }

        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="FileName">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static Config Load(string FileName = null)
        {


            var fileName = FileName == null ?
                DefaultFileName ://no filename use default config
                FileName;

            var fi = new FileInfo(fileName);

            if (!fi.Exists && FileName == null) return new Config();

            using (var stream = File.OpenRead(fileName))
            {
                var serializer = new XmlSerializer(typeof(Config));
                return serializer.Deserialize(stream) as Config;
            }

        }

        public virtual object Clone()
        {
            MemoryStream ms = new MemoryStream();
            BinaryFormatter bf = new BinaryFormatter();

            bf.Serialize(ms, this);

            ms.Position = 0;
            object obj = bf.Deserialize(ms);
            ms.Close();

            return obj as Config;
        }

        #endregion method
    }
}
