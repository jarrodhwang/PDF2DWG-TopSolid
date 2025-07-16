using FormDialog.Dialog;
using PDF2DWG.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Automation.lib.Pdm.type;
using TopSolid.Kernel.Automating;
using Automation.lib.Pdm;
using System.IO;

namespace PDF2DWG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Hide();

         
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            if(SelectFile())
               if( PDF2DWG())
                    DWG2TopSolid();

            this.Close();
        }

        string inFile = "";
        string outFile = "";
        private bool SelectFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            { 
                ofd.CheckFileExists = true;
                ofd.Multiselect = false; // 배치 변환 
                ofd.Filter = "PDF 파일 (*.pdf)|*.pdf|모든 파일 (*.*)|*.*";
                ofd.Title = "3DPDF 파일 선택";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    
                    inFile = ofd.FileName;
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        
        private bool PDF2DWG()
        {
            // 1. 실행 파일 경로와 파라미터 지정
            string exePath = @"C:\Program Files (x86)\Any PDF to DWG Converter\pdf_dwg.exe";

            outFile = Path.ChangeExtension(inFile, ".dwg");
            string args = $"/InFile \"{inFile}\" /OutFile \"{outFile}\" /Hide /Overwrite";


            // 2. ProcessStartInfo 설정
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args,
                UseShellExecute = false,       // 표준 입출력 리디렉션을 원할 때 false
                RedirectStandardOutput = true, // 결과를 코드로 받으려면 true
                RedirectStandardError = true,
                CreateNoWindow = true          // 창 안 띄움
            };


            // 3. 실행
            try
            {
                using (var proc = Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    // 결과(로그) 처리


                    if (!string.IsNullOrEmpty(error))
                        CallErrorDialog(error, true);

                    if (proc.ExitCode == 0)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                CallErrorDialog("pdf to dwg error"+ex, true);
                return false;
            }
            return false;
        }

        private void DWG2TopSolid()
        {

            var importFileToPrj = CallProjectDialog(new List<DocumentType>() { DocumentType.Project, DocumentType.Folder }, false);
            if (importFileToPrj.DialogResult == DialogResult.Cancel)
                return;
            var selectedDoc = importFileToPrj.itemToSend;
            

            try
            {

                // 비동기로 파일 변환하여 가져오기
                using (var loadingDialog = new LoadingDialog("변환해서 가져오는 중", 300 ,300,  25) { Owner = this, StartPosition = FormStartPosition.CenterScreen, Visible = false})
                {
                    loadingDialog.Show(this);
                    //선택한 위치에 가져오기


                    //var importOptions = new List<KeyValue>();
                    //if (true)
                    //    importOptions.Add(new KeyValue("ASSEMBLY_DOCUMENT_EXTENSION", ".TopPrt"));
                    //if (true)
                    //    importOptions.Add(new KeyValue("SIMPLIFIES_GEOMETRY", "False"));
                    //if (true)
                    //    importOptions.Add(new KeyValue("TRANSLATES_ASSEMBLY", "True"));

                    DocumentId import = new DocumentId();
                    
                    //if (importOptions.Count > 0)
                    //{
                    //    import = ProjectManager.ImportFile(inFile, selectedDoc, importOptions);
                    //}
                    //else
                    //{
                        import = ProjectManager.ImportFile(outFile, selectedDoc);
                    //}


                    selectedDoc = TopSolidHost.Documents.GetPdmObject(import);
                    //selectedDocument = selectedDoc;
                    TopSolidHost.Pdm.Save(selectedDoc, false);
                    loadingDialog.Close();


                    //DisplayPartButton(this.selectPartButton, selectedDoc);

                }

                CallErrorDialog("변환 완료", false);


            }
            catch (InvalidOperationException) { }
            catch
            {
                CallErrorDialog("가져오기 실패함", true);
            }

        }

        public ProjectDialog CallProjectDialog(List<DocumentType> inDocTypes, bool isLibrary)
        {
            ProjectDialog prj = new ProjectDialog(inDocTypes, isLibrary)
            {
                Owner = this, // 이 폼의 소유자를 메인폼으로
                StartPosition = FormStartPosition.CenterScreen // 폼 생성 위치를 메인폼의 중앙으로
            };
            prj.ShowDialog();

            return prj;
        }

        public InfoDialog CallErrorDialog(string errorMsg, bool isError)
        {
            var errorDialog = new InfoDialog(errorMsg, isError)
            {
                Owner = this,
                StartPosition = FormStartPosition.CenterScreen,
                TopMost = true
            };
            errorDialog.ShowDialog();

            return errorDialog;
        }

       
    }
}
