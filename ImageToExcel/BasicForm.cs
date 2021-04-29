using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Net.Http;
using System.Linq;
using Newtonsoft.Json;

namespace ImageToExcel
{
    /// <summary>
    /// Basic form
    /// </summary>
    /// <seealso cref="System.Windows.Forms.Form" />
    public partial class BasicForm : Form
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="BasicForm"/> class.
        /// </summary>
        public BasicForm()
        {
            this.InitializeComponent();
        }

        #endregion

        #region Events

        /// <summary>
        /// Handles the Click event of the ConvertToText button control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private async void btnConvertToText_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = string.Empty;
                //// https://www.cmarix.com/git/DotNet/OCR-ImageToText-Example/Sample_1.tiff
                //AutoOcr autoOcr = new AutoOcr();
                string FolderPath = tbInputURL.Text;
                List<string> inputImagesPath = this.GetImageFromFolder(FolderPath);
                if (inputImagesPath == null)
                {
                    lblStatus.Text = "Error : Invalid Path";
                    return;
                }
                else if (inputImagesPath.Count == 0)
                {
                    lblStatus.Text = "Error : Images Don't exist";
                    return;
                }

                string outputfile = FolderPath + "/output/" + DateTime.Now.ToString("yyyy_mm_dd__HH_mm_ss") + ".xlsx";

                List<string> outputResultText = new List<string>();
                int index = 1;
                foreach (string ImagePath in inputImagesPath)
                {
                    if (index > 1)
                    {
                        break;
                    }

                    if (string.IsNullOrEmpty(ImagePath))
                        return;

                    string allText = "";

                    try
                    {
                        HttpClient httpClient = new HttpClient();
                        httpClient.Timeout = new TimeSpan(1, 1, 1);


                        MultipartFormDataContent form = new MultipartFormDataContent();
                        form.Add(new StringContent("96bc3f612088957"), "apikey"); //Added api key in form data
                        form.Add(new StringContent("eng"), "language");

                        form.Add(new StringContent("2"), "ocrengine");
                        form.Add(new StringContent("true"), "scale");
                        form.Add(new StringContent("true"), "istable");

                        HttpResponseMessage response = await httpClient.PostAsync("https://api.ocr.space/Parse/Image", form);

                        string strContent = await response.Content.ReadAsStringAsync();



                        Rootobject ocrResult = JsonConvert.DeserializeObject<Rootobject>(strContent);

                        if (ocrResult.OCRExitCode == 1)
                        {
                            for (int i = 0; i < ocrResult.ParsedResults.Count(); i++)
                            {
                                allText = allText + ocrResult.ParsedResults[i].ParsedText;
                            }
                        }
                        else
                        {
                            MessageBox.Show("ERROR: " + strContent);
                        }
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show("Ooops");
                    }

                    outputResultText = analysisText(allText);

                    index++;
                }


                ProcessToExcel(outputfile, outputResultText);

                lblStatus.Text = "Finish : Folder name -" + tbInputURL.Text;
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Error : " + ex.Message;
            }
        }

        #endregion

        public List<string> analysisText(string text)
        {
            List<string> result = new List<string>();
            result.Add("EVOCARE");
            result.Add("C/O Accounts Payable");
            result.Add("4735 Melissa way");
            result.Add("Birmingham");
            result.Add("Al");
            result.Add("35243");
            return result;
        }

        #region write the date to Excel file
        private void ProcessToExcel(string outputFile, List<string> datas)
        {
            List<string> headers = new List<string>();
            headers.Add("payor name");
            headers.Add("payor atten");
            headers.Add("Payor street address");
            headers.Add("Payor city");
            headers.Add("Payor state");
            headers.Add("Payor zip");
            headers.Add("1) Plan type");
            headers.Add("1a) Insured's ID number");
            headers.Add("2) Patient last name");
            headers.Add("2) Patient first name");
            headers.Add("2) Patient Middle");
            headers.Add("3) Patient DOB");
            headers.Add("3) Patient sex");
            headers.Add("5) Patient street address");
            headers.Add("5) Patient city");
            headers.Add("5) Patient state");
            headers.Add("5) Patient zip");
            headers.Add("5) Patient phone number");
            headers.Add("6) Patient relationship to insured");
            headers.Add("4) Insured last name");
            headers.Add("4) insured first name");
            headers.Add("4) insured middle");
            headers.Add("7) Insured street address");
            headers.Add("7) Insured city");
            headers.Add("7) Insured state");
            headers.Add("7) Insured zip");
            headers.Add("11) Insured policy number");
            headers.Add("11a) insured date of birth (DOB)");
            headers.Add("11c) Insurance plan name");
            headers.Add("11d) Is there another health plan");
            headers.Add("10a) Patient condition related to Employment");
            headers.Add("10b) Patient condition related to auto accident");
            headers.Add("10c) Patient condition related to other accident");
            headers.Add("12) authorized person");
            headers.Add("12) Date of authorization");
            headers.Add("13) insured authorzation");
            headers.Add("17 Name of referring physician");
            headers.Add("17a)");
            headers.Add("17b)");
            headers.Add("20) outside lab");
            headers.Add("21a) diagnosis");
            headers.Add("21b) diagnosis");
            headers.Add("ICD indicator");
            headers.Add("24a) dates of service");
            headers.Add("24b) place of service");
            headers.Add("24d) procedure service");
            headers.Add("24d) procedure service modifier");
            headers.Add("24e) proceudre pointer");
            headers.Add("24f) procedure charges");
            headers.Add("24g) procedure days or units");
            headers.Add("24j) procedure provider NPI");
            headers.Add("25) Federal tax ID");
            headers.Add("25) SSN or EIN");
            headers.Add("26 patient's account #");
            headers.Add("27) accept assignment");
            headers.Add("28) Total charges");
            headers.Add("31) signature of physician");
            headers.Add("31) date of signature");
            headers.Add("32) service facility name");
            headers.Add("32) service facility street address");
            headers.Add("32) service facility city");
            headers.Add("32) service facility street state");
            headers.Add("32) service facility street zip");
            headers.Add("33) billing provider phone number");
            headers.Add("33) billing provider name");
            headers.Add("33) billing provider street address");
            headers.Add("33) billing provider city");
            headers.Add("33) billing provider State");
            headers.Add("33) billing provider zip");
            headers.Add("33a)");
            headers.Add("33b)");

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;

            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                for (int i = 1; i <= headers.Count; i++)
                {
                    oSheet.Cells[1, i] = headers[i - 1];
                }

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "BS1").Font.Bold = true;


                //analysis the data
                int nRow = 2;
                for (int i = 1; i <= datas.Count; i++)
                {
                    oSheet.Cells[nRow, i] = datas[i - 1];
                }

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.Saved = true;
                oWB.SaveCopyAs(outputFile);
                oWB.Close(true, outputFile, Type.Missing);

                oXL.Quit();
            }
            catch (Exception Ex)
            {
                string strErr = Ex.Message;
            }

            return;
        }

        #region Helper Methods
        private List<string> GetImageFromFolder(string FolderPath)
        {
            Boolean isFolderExists = false;
            isFolderExists = Directory.Exists(FolderPath);
            List<string> imageResult = new List<string>();
            if (!isFolderExists)
                return null;

            string[] files = Directory.GetFiles(FolderPath);

            string outputdirectory = FolderPath + "/output/";
            if (!Directory.Exists(outputdirectory))
            {
                Directory.CreateDirectory(outputdirectory);
            }

            foreach (string filepath in files)
            {
                imageResult.Add(filepath);
            }

            return imageResult;
        }

        #endregion
    }
}
#endregion