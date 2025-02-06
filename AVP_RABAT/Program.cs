using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using SMBLibrary.Client;
using SMBLibrary;

namespace AVP_RABAT
{
    internal class Program
    {
        static Dictionary<int, float> vwRabat = new Dictionary<int, float>();
        static Dictionary<int, float> skodaRabat = new Dictionary<int, float>();
        static Dictionary<int, float> seatRabat = new Dictionary<int, float>();
        static Dictionary<int, float> porscheRabat = new Dictionary<int, float>();

        static List<AvpData> vwData = new List<AvpData>();
        static List<AvpData> skodaData = new List<AvpData>();
        static List<AvpData> seatData = new List<AvpData>();
        static List<AvpData> porscheData = new List<AvpData>();

        static string rabatFile = @"\\Fs\arlista\Árlista 2025\GYÁRI ÁRLISTÁK 2025\AVP_RABAT_NEW.xlsx";
        static string vwFile = @"C:\Users\LP-KATALOGUS1\source\repos\AVP_RABAT\AVP_RABAT\bin\Debug\kezi\vwFull.csv";
        static string skodaFile = @"\\Fs\arlista\kezi_arlista\skoda_avp.csv";
        static string seatFile = @"\\Fs\arlista\kezi_arlista\seat_avp.csv";
        static string porscheFile = @"\\Fs\arlista\kezi_arlista\porsche_avp.csv";
        static void Main(string[] args)
        {
            LoadRabatInfos();

            /*vwData = EvaluateAvpFile(vwFile);
            skodaData = EvaluateAvpFile(skodaFile);
            porscheData = EvaluateAvpFile(porscheFile);
            seatData = EvaluateAvpFile(seatFile);*/

            EditFile(vwRabat, vwFile, "vw_avp.csv", @"kezi_arlista\vw_avp.csv");
            //EditFile(skodaRabat, skodaFile, "skoda_avp.csv", @"kezi_arlista\skoda_avp.csv");
            //EditFile(seatRabat, seatFile, "seat_avp.csv", @"kezi_arlista\seat_avp.csv");
            //EditFile(porscheRabat, porscheFile, "porsche_avp.csv");
        }

        static void EditFile(Dictionary<int, float> dict, string file, string fileName, string smbPath)
        {
            Console.WriteLine(fileName);
            StreamReader sr = new StreamReader(file);
            StreamWriter sw = new StreamWriter(fileName, false, sr.CurrentEncoding);
            Console.WriteLine(sr.CurrentEncoding);
            string header = sr.ReadLine();
            string[] headerData = header.Split(';');
            int priceCol = GetPriceCol(headerData);
            int rabatCol = GetRabatCol(headerData);
            int finalPriceCol = GetFinalPriceCol(headerData);

            sw.WriteLine(header);

            if (priceCol == -1 || rabatCol == -1 || finalPriceCol == -1)
            {
                Console.WriteLine("Hibás fájl {0}", fileName);
                return;
            }

            while (!sr.EndOfStream)
            {
                string[] line = sr.ReadLine().Split(';');
                float price = float.Parse(line[priceCol]);
                if (dict.ContainsKey(int.Parse(line[rabatCol])))
                    line[finalPriceCol] = (price - price * dict[int.Parse(line[rabatCol])]).ToString();

                sw.WriteLine(string.Join(";", line));
            }
            sw.Close();
            sr.Close();

            //SaveToSmb(smbPath, fileName);
        }

        static List<AvpData> EvaluateAvpFile(string file)
        {
            Console.WriteLine(file);
            List<AvpData> datas = new List<AvpData>();
            StreamReader sr = new StreamReader(file);
            string header = sr.ReadLine();
            string[] headerData = header.Split(';');
            int priceCol = GetPriceCol(headerData);
            int rabatCol = GetRabatCol(headerData);

            if (priceCol == -1 || rabatCol == -1)
            {
                Console.WriteLine("Hibás fájl {0}", file);
                return null;
            }

            while (!sr.EndOfStream)
            {
                string[] line = sr.ReadLine().Split(';');
                AvpData data = new AvpData();
                data.rabat = int.Parse(line[rabatCol]);
                data.price = float.Parse(line[priceCol]);
                datas.Add(data);
            }

            sr.Close();

            return datas;
        }

        static int GetPriceCol(string[] firstLineData)
        {
            for (int i = 0; i < firstLineData.Length; i++)
            {
                if (firstLineData[i].Trim() == "Preis") return i;
            }
            return -1;
        }

        static int GetRabatCol(string[] firstLineData)
        {
            for (int i = 0; i < firstLineData.Length; i++)
            {
                if (firstLineData[i].Trim() == "RG") return i;
            }
            return -1;
        }

        static int GetFinalPriceCol(string[] firstLineData)
        {
            for (int i = 0; i < firstLineData.Length; i++)
            {
                if (firstLineData[i].Trim() == "price") return i;
            }
            return -1;
        }

        static void LoadRabatInfos()
        {
            var wb = new XLWorkbook(rabatFile);
            var ws = wb.Worksheet("Munka1");
            vwRabat = GetColumnData(ws, 1);
            skodaRabat = GetColumnData(ws, 4);
            seatRabat = GetColumnData(ws, 8);
            porscheRabat = GetColumnData(ws, 11);
        }

        static Dictionary<int, float> GetColumnData(IXLWorksheet ws, int column)
        {
            Dictionary<int, float> data = new Dictionary<int, float>();

            var col = ws.Column(column);
            var colP = ws.Column(column + 1);
            int i = 3;
            while (!col.Cell(i).IsEmpty())
            {
                data.Add(int.Parse(col.Cell(i).Value.ToString()), float.Parse(colP.Cell(i).Value.ToString()));
                i++;
            }

            return data;
        }

        private static void SaveToSmb(string savePath, string fileName)
        {
            //This function works on black magic
            SMB2Client client = new SMB2Client();
            bool isConnected = client.Connect(IPAddress.Parse("131.0.2.20"), SMBTransportType.DirectTCPTransport);
            if (isConnected)
            {
                NTStatus statuss = client.Login(String.Empty, "brobert", "ma1beSu5");
                if (statuss == NTStatus.STATUS_SUCCESS)
                {
                    Console.WriteLine("Sikeres belépés");
                }
            }

            ISMBFileStore fileStore = client.TreeConnect(@"arlista", out NTStatus stat);
            if (fileStore is SMB2FileStore == false)
                return;
            object fileHandle;
            FileStatus fileStatus;
            NTStatus status = fileStore.CreateFile(out fileHandle, out fileStatus, savePath, AccessMask.GENERIC_WRITE | AccessMask.DELETE | AccessMask.SYNCHRONIZE, SMBLibrary.FileAttributes.Normal, ShareAccess.None, CreateDisposition.FILE_OPEN, CreateOptions.FILE_NON_DIRECTORY_FILE | CreateOptions.FILE_SYNCHRONOUS_IO_ALERT, null);

            //Turns out, there is no owerwrite in this SMB library.
            //So we first delete the old file
            if (status == NTStatus.STATUS_SUCCESS)
            {
                FileDispositionInformation fileDispositionInformation = new FileDispositionInformation();
                fileDispositionInformation.DeletePending = true;
                status = fileStore.SetFileInformation(fileHandle, fileDispositionInformation);
                bool deleteSucceeded = (status == NTStatus.STATUS_SUCCESS);
                status = fileStore.CloseFile(fileHandle);
            }

            //We should be making a backup, "old file", but it's a hassle just to save one file
            //If only I had more time to work on this
            status = fileStore.CreateFile(out fileHandle, out fileStatus, savePath, AccessMask.GENERIC_WRITE | AccessMask.SYNCHRONIZE, SMBLibrary.FileAttributes.Normal, ShareAccess.None, CreateDisposition.FILE_CREATE, CreateOptions.FILE_NON_DIRECTORY_FILE | CreateOptions.FILE_SYNCHRONOUS_IO_ALERT, null);
            if (status == NTStatus.STATUS_SUCCESS)
            {
                int bitesWritten;
                byte[] bites = File.ReadAllBytes(fileName);
                status = fileStore.WriteFile(out bitesWritten, fileHandle, 0, bites);
                if (status != NTStatus.STATUS_SUCCESS)
                {
                    throw new Exception("Failed to write to a file!");
                }
                status = fileStore.CloseFile(fileHandle);
            }
            status = fileStore.Disconnect();
            Console.WriteLine("SMB File Writing Complete");
        }
    }

    public struct AvpData
    {
        public int rabat { get; set; }
        public float price { get; set; }
    }
}
