// using Meaghan_HCPS_Dev data and check duplicates on completes test,not on file name.
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ImportUsingRadioButton
{
    /// <summary>
    /// This Form is used to Import Data from .dat (data files) and create Excel file for each ClassRoom and each group Id.
    /// </summary>
    public partial class Form1 : Form
    {
        //Global Variable which is used to get Machine User Name.
        string _userName = Environment.UserName;
        Excel.Application excelApp = new Excel.Application();
        //Global Variable to store user selected value from Form. Intializing it to Empty string.
        string _rootfolder = string.Empty;

        //Global Variables to store paths, which will use to read .dat files and roster files.
        string _commonPath = @"C:\Users\";
        string _boxDrivePath = @"\Box\Meagan_HCPS\Participants";
        string _boxsyncPath = @"\Box Sync\Meagan_HCPS\Participants";

        //This  Collections will hold all the classroom IDs.
        static int[] _classRoomIds = new int[] { 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 };

        //This Collection will hold all time points. 
        //int[] _timePoints = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        int timePoint;
        String finalTestName = "ELSMC";
        //This Collection will be used to Create Output Excel file.
        private List<OutputDto> _result = new List<OutputDto>();
         string currentTime = DateTime.Now.ToString("yyyyMMddHHmmss");
        Boolean errorFlag = false;
        string rootFolderforDatfiles= "";
        string logFilePath1 = "";
        //  ArrayList files = new ArrayList();
        List<int> logEntry = new List<int>();
        ArrayList duplicateList = new ArrayList();

        static int classID = 10;
        static bool logFlag = true;
        static bool Flag = true;
        /// <summary>
        /// Constructor, which loads all intial Data.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// This Method is Executed on Button Click Event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //This gets called when user clicks Import button
            button1.Enabled = false;
            Import(_rootfolder);



            if (errorFlag == false) // No duplicates found in input files
            {
                if (MessageBox.Show(this, "Import completed without any errors.Do you want to exit application?", "SRCBM_ImportSpreadsheets", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                   Application.Exit();

                else
                    button1.Enabled = true;

            }

            else

            {
                StringBuilder logpath = new StringBuilder();
                logpath.Append(rootFolderforDatfiles);
                logpath.Append(@"\Inquisit_data\Log");

                string msg = "Import completed and found Duplicate IDs. Do you want to open the log File?";
                if(MessageBox.Show(this,msg, "SRCBM_ImportSpreadsheets",MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (Directory.Exists(logpath.ToString()))
                    {
                  
                        Process.Start(logFilePath1);
                        Application.Exit();
                    }
                    else
                    {
                        Application.Exit();
                    }

                }

                else
                    button1.Enabled = true;

            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BoxSync_CheckedChanged(object sender, EventArgs e)
        {
            //Set the rootfolder path to boxsync if user has selected BoxSync
            if (BoxSync.Checked == true)
                _rootfolder = PreparePath(_commonPath, _boxsyncPath,true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //Set the rootfolder path to boxdrive if user has selected BoxDrive
              if (BoxDrive.Checked == true)
                    _rootfolder = PreparePath(_commonPath, _boxDrivePath,true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commonPath"></param>
        /// <param name="specificPath"></param>
        /// <returns></returns>
        private string PreparePath(string commonPath, string specificPath,bool includeUsername)
        {
            StringBuilder s1 = new StringBuilder();
            s1.Append(commonPath);
            if (includeUsername)
            {
                s1.Append(_userName);
            }
            s1.Append(specificPath);

            return s1.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        private void ReadExcel(string fileName, int id, int group = 0, bool byRoster = false)
        {
            
            FileStream stream = null;
            IExcelDataReader excelReader = null;
            System.Data.DataTable dataTable = new DataTable();
            
                stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            
            
            List<string> results = new List<string>();
            ArrayList duplicateList = new ArrayList();

            using (stream)
            {
                if (byRoster)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream); // to read roster files
                else
                    excelReader = ExcelReaderFactory.CreateCsvReader(stream);    // to read  .dat files

                var result = excelReader.AsDataSet(new ExcelDataSetConfiguration() // NUGET package___Converts Excel file into dataset(Datatable)
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() //its just a configuration to read excel
                    {
                        UseHeaderRow = true
                    }
                }); // Now result has whole excel data

                if (result != null && result.Tables.Count > 0) // atleast one sheet should be present in Excel
                {
                    if (byRoster == true)
                        dataTable = result.Tables[0].Select().CopyToDataTable(); // take sheet1 data
                    else
                    {
                        var rows = result.Tables[0].AsEnumerable().Where(x => x.Field<string>("group") == group.ToString());
                        dataTable = rows.Any() ? rows.CopyToDataTable() : dataTable;
                        // try to collect each test data together.
                            
                        duplicateList =IdentifyDuplicateRows(dataTable,"subject",fileName,id);

                    }
                }
            }

            foreach (DataRow dr in dataTable.Rows)
            {
                if (fileName.Contains("Roster"))
                {
                    OutputDto dto = new OutputDto();
                    dto.SubJectId = (dr["ID"]).ToString(); //data row of ID column
                    _result.Add(dto);
                }

                if (fileName.Contains("ELNMC"))
                {

                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test1Score = (dr["values.total_correct"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test1Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("ELNFR"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test2Score = (dr["values.total_correct"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test2Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("ELSMC"))
                {

                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test3AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test3BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test3Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("ELSFR"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test4AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test4BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test4Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("ERHYM"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test5AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test5BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test5Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("EBLMC"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test6AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test6BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test6Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("EBLFR"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test7AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test7BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test7Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }

                if (fileName.Contains("EVOCB"))
                {
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test8AScore = (dr["values.Item_count_correctA"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test8BScore = (dr["values.Item_count_correctB"]).ToString(); return a; }).ToList();
                    _result.Where(i => (i.SubJectId == (dr["subject"]).ToString()) && (duplicateList.Contains(dr["subject"]) == false)).Select(a => { a.Test8Date = DateTime.ParseExact(dr["date"].ToString(), "MMddyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy"); return a; }).ToList();
                }
            }
           
        }

        private ArrayList IdentifyDuplicateRows(DataTable dTable, string colName, string testName, int id)
        {
            Hashtable hTable = new Hashtable();
            string spaces = "";
            string classroom = "";
            string number = "";



            if (classID == 10 && logFlag == true)
            {
                spaces = "===============================";
                classroom = "Class ID:";
                classID = id;
                number = id.ToString();
                logFlag = false;
            }


            else if (classID != id)
            {
                spaces = "===============================";
                classroom = "Class ID:";
                classID = id;
                number = id.ToString();
            }


            string sub = "HCPS";
            int index = testName.LastIndexOf(sub);

            string finalTestName11 = testName.Substring(index + 5, 5);


           
            string rootFolderforLogfile = _rootfolder.Substring(0, _rootfolder.Length - 13);

            StringBuilder logFilePath = new StringBuilder();
            logFilePath.Append(rootFolderforLogfile);
            logFilePath.Append(@"\Inquisit_data\Log\SRCBM_Error_Log_");
            //logFilePath.Append(id);
            //logFilePath.Append("_");
            logFilePath.Append(timePoint);
            logFilePath.Append("_");
            logFilePath.Append(currentTime);
            logFilePath.Append(".txt");

            logFilePath1 = logFilePath.ToString();



            //  String finalTestName1 = testName.Substring(index1 + 10);
            if (finalTestName == "ELSMC" && Flag == true)
            {

                using (StreamWriter writer = new StreamWriter(logFilePath.ToString(), true))
                {
                    writer.WriteLine(spaces + classroom + number + spaces);
                    finalTestName = finalTestName11;
                    Flag = false;
                }
            }

         
            //=================

          else  if (finalTestName != finalTestName11)
            {
                try
                {
                    using (StreamWriter writer = new StreamWriter(logFilePath.ToString(), true))
                    {
                        writer.WriteLine();
                        writer.WriteLine(finalTestName);

                        if (duplicateList.Count == 0)
                        {
                            writer.Write("No Errors found in Records");
                            writer.Write("");
                        }
                        else
                        {
                            writer.Write("Duplicate IDs found in records:");
                            errorFlag = true;
                            foreach (var entry in duplicateList)
                            {
                                writer.Write(entry);
                                writer.Write(" ");
                            }
                            duplicateList.Clear();
                        }
                        writer.WriteLine();
                        writer.WriteLine(spaces + classroom + number + spaces);                        
                        finalTestName = finalTestName11;
                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                

            }

            else

            foreach (DataRow drow in dTable.Rows)
            {
                if (hTable.Contains(drow[colName]))
                {
                    string childId = drow["subject"].ToString();

                    if(childId.Length == 4)
                        {
                            if (String.Compare(childId.Substring(0, 2), id.ToString()) == 0)
                                duplicateList.Add(drow["subject"]);

                        }
                    
                }
                else
                    hTable.Add(drow[colName], string.Empty);
            }
            
            return duplicateList;
        }


        /// <summary>
        /// This Method is used to Read each classroom and select each group id and create Excel File.
        /// </summary>
        /// <param name="rootFolder"></param>
        private void Import(string rootFolder)
        {
           
            rootFolderforDatfiles = _rootfolder.Substring(0, _rootfolder.Length - 13);
          

            //Iterate through each classroom in _classRoomIds, to create Output Excel file for each ID.
            foreach (var id in _classRoomIds)
            {
                //string builder used to create input path for reading files.
                StringBuilder rosterFilePath = new StringBuilder(rootFolder);

                rosterFilePath.Append("\\SRCBM_Roster_");
                rosterFilePath.Append(id);
                rosterFilePath.Append(".xlsx");
                try
                {                   
                    //Read Roster File.
                    ReadExcel(rosterFilePath.ToString(), id, 0, true);

                        //Read Input files for .dat.
                        CreateInputFilePaths(timePoint,id);

                        //Create Ouput Excel for each time point and classroom.
                        CreateOutputExcelFile(id, timePoint);

                        //Intialize result list for new Classroom
                        _result = new List<OutputDto>();
                        
                }
                catch (Exception Ex)
                {
                    
                    Console.WriteLine(Ex.ToString());
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="classroomId"></param>
        /// <param name="timepoint"></param>
        private void CreateOutputExcelFile(int classroomId, int timePoint)
        {
            //var excelApp = new Excel.Application();

            // Make the object visible.
            excelApp.Visible = false;

            Microsoft.Office.Interop.Excel.Worksheet workSheet;
            Microsoft.Office.Interop.Excel.Workbook wbook;

            wbook = excelApp.Workbooks.Add(true);
            workSheet = wbook.ActiveSheet;

            // Establish column headings in cells
            workSheet.Cells[1, "A"] = "ID";
            workSheet.Cells[1, "B"] = "ELNMC_date";
            workSheet.Cells[1, "C"] = "ELNMC_Score";

            workSheet.Cells[1, "D"] = "ELNFR_date";
            workSheet.Cells[1, "E"] = "ELNFR_Score";

            workSheet.Cells[1, "F"] = "ELSMC_date";
            workSheet.Cells[1, "G"] = "ELSMCA_score";
            workSheet.Cells[1, "H"] = "ELSMCB_score";

            workSheet.Cells[1, "I"] = "ELSFR_date";
            workSheet.Cells[1, "J"] = "ELSFRA_score";
            workSheet.Cells[1, "K"] = "ELSFRB_score";

            workSheet.Cells[1, "L"] = "ERHYM_date";
            workSheet.Cells[1, "M"] = "RHYMA_score";
            workSheet.Cells[1, "N"] = "RHYMB_score";

            workSheet.Cells[1, "O"] = "EBLMC_date";
            workSheet.Cells[1, "P"] = "EBLMCA_Score";
            workSheet.Cells[1, "Q"] = "EBLMCB_Score";

            workSheet.Cells[1, "R"] = "EBLFR_date";
            workSheet.Cells[1, "S"] = "EBLFRA_Score";
            workSheet.Cells[1, "T"] = "EBLFRB_Score";

            workSheet.Cells[1, "U"] = "EVOC_date";
            workSheet.Cells[1, "V"] = "EVOCA_Score";
            workSheet.Cells[1, "W"] = "EVOCB_Score";

            int row = 1;
            foreach (var dto in _result)
            {
                row++;
                workSheet.Cells[row, "A"] = dto.SubJectId;
                workSheet.Cells[row, "B"] = dto.Test1Date;
                workSheet.Cells[row, "C"] = dto.Test1Score;

                workSheet.Cells[row, "D"] = dto.Test2Date;
                workSheet.Cells[row, "E"] = dto.Test2Score;

                workSheet.Cells[row, "F"] = dto.Test3Date;
                workSheet.Cells[row, "G"] = dto.Test3AScore;
                workSheet.Cells[row, "H"] = dto.Test3BScore;

                workSheet.Cells[row, "I"] = dto.Test4Date;
                workSheet.Cells[row, "J"] = dto.Test4AScore;
                workSheet.Cells[row, "K"] = dto.Test4BScore;

                workSheet.Cells[row, "L"] = dto.Test5Date;
                workSheet.Cells[row, "M"] = dto.Test5AScore;
                workSheet.Cells[row, "N"] = dto.Test5BScore;

                workSheet.Cells[row, "O"] = dto.Test6Date;
                workSheet.Cells[row, "P"] = dto.Test6AScore;
                workSheet.Cells[row, "Q"] = dto.Test6BScore;

                workSheet.Cells[row, "R"] = dto.Test7Date;
                workSheet.Cells[row, "S"] = dto.Test7AScore;
                workSheet.Cells[row, "T"] = dto.Test7BScore;

                workSheet.Cells[row, "U"] = dto.Test8Date;
                workSheet.Cells[row, "V"] = dto.Test8AScore;
                workSheet.Cells[row, "W"] = dto.Test8BScore;
            }

            //Make All cells AutoFit
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
            workSheet.Columns[5].AutoFit();
            workSheet.Columns[6].AutoFit();
            workSheet.Columns[7].AutoFit();
            workSheet.Columns[8].AutoFit();
            workSheet.Columns[9].AutoFit();
            workSheet.Columns[10].AutoFit();
            workSheet.Columns[11].AutoFit();
            workSheet.Columns[12].AutoFit();
            workSheet.Columns[13].AutoFit();
            workSheet.Columns[14].AutoFit();
            workSheet.Columns[15].AutoFit();
            workSheet.Columns[16].AutoFit();
            workSheet.Columns[17].AutoFit();
            workSheet.Columns[18].AutoFit();
            workSheet.Columns[19].AutoFit();
            workSheet.Columns[20].AutoFit();
            workSheet.Columns[21].AutoFit();
            workSheet.Columns[22].AutoFit();
            workSheet.Columns[23].AutoFit();

            //File name Format: SRCBM_import_classroomID_timepoint
            // if file already exists with same filename then move to backup folder
            // Create backup folder, check all filenames from current directory, if same file exist then move to destination path. 

            string rootFolderforDatfiles = _rootfolder.Substring(0, _rootfolder.Length - 13);

            string backUpfolderPath = PreparePath(rootFolderforDatfiles, @"\Inquisit_data\Backup",false);

            StringBuilder filePathtoDelete = new StringBuilder();
            filePathtoDelete.Append(rootFolderforDatfiles);
            filePathtoDelete.Append(@"\Inquisit_data\SRCBM_Import_");
            filePathtoDelete.Append(classroomId);
            filePathtoDelete.Append("_");
            filePathtoDelete.Append(timePoint);
            filePathtoDelete.Append(".xlsx");

            
           

            StringBuilder destinationFile = new StringBuilder();
            destinationFile.Append(rootFolderforDatfiles);
            destinationFile.Append(@"\Inquisit_data\Backup\SRCBM_Import_");
            destinationFile.Append(classroomId);
            destinationFile.Append("_");
            destinationFile.Append(timePoint);
            destinationFile.Append("_");
            destinationFile.Append(currentTime);
            destinationFile.Append(".xlsx");           

            System.IO.Directory.CreateDirectory(backUpfolderPath.ToString());           

            if (File.Exists(filePathtoDelete.ToString()))
                File.Move(filePathtoDelete.ToString(), destinationFile.ToString());

            wbook.SaveAs(filePathtoDelete.ToString(), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
             false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wbook.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="timePoint"></param>
        private void CreateInputFilePaths(int timePoint, int id)
        {
            string[] inputfilePaths = new string[8];
            List<string[]> files = new List<string[]>();
            rootFolderforDatfiles = _rootfolder.Substring(0, _rootfolder.Length - 13);
            //int id1 = id;
            //production data
            //inputfilePaths[0] = PreparePath(rootFolderforDatfiles, @"\ELSMC\ELSMC_Validation_ShortForms_AB_v1.dat", false);
            //inputfilePaths[1] = PreparePath(rootFolderforDatfiles, @"\EBLFR\EBLFR_SHORTFORMSAB_Final.dat", false);
            //inputfilePaths[2] = PreparePath(rootFolderforDatfiles, @"\EBLMC\EBLMC_SHORTFORMSAB_Final.dat", false);
            //inputfilePaths[3] = PreparePath(rootFolderforDatfiles, @"\ELSFR\ELSFR_Validation_ShortFormAB_v1.dat", false);

            //inputfilePaths[4] = PreparePath(rootFolderforDatfiles, @"\ELNMC\ELNMC_SRCBM_Validation_ShortForms_Final.dat", false);
            //inputfilePaths[5] = PreparePath(rootFolderforDatfiles, @"\EVOCB\EVOCB_SHORTFORMSAB_Final .dat", false);
            //inputfilePaths[6] = PreparePath(rootFolderforDatfiles, @"\ELNFR\ELNFR_SRCBM_Validation_ShortForms_V2.dat", false);
            //inputfilePaths[7] = PreparePath(rootFolderforDatfiles, @"\ERHYM\ERHYM_Validation_ShortFormAB_v1.dat", false);

            ////test data
            //inputfilePaths[0] = PreparePath(rootFolderforDatfiles, @"\ELSMC\test_data\ELSMC_Validation_ShortForms_AB_v1.dat", false);
            //inputfilePaths[1] = PreparePath(rootFolderforDatfiles, @"\EBLFR\test_data\EBLFR_SHORTFORMSAB_Final.dat", false);
            //inputfilePaths[2] = PreparePath(rootFolderforDatfiles, @"\EBLMC\test_data\EBLMC_SHORTFORMSAB_Final.dat", false);
            //inputfilePaths[3] = PreparePath(rootFolderforDatfiles, @"\ELSFR\test_data\ELSFR_Validation_ShortFormAB_v1.dat", false);

            //inputfilePaths[4] = PreparePath(rootFolderforDatfiles, @"\ELNMC\test_data\ELNMC_SRCBM_Validation_ShortForms_Final.dat", false);
            //inputfilePaths[5] = PreparePath(rootFolderforDatfiles, @"\EVOCB\test_data\EVOCB_SHORTFORMSAB_Final.dat", false);
            //inputfilePaths[6] = PreparePath(rootFolderforDatfiles, @"\ELNFR\test_data\ELNFR_SRCBM_Validation_ShortForms_V2.dat", false);
            //inputfilePaths[7] = PreparePath(rootFolderforDatfiles, @"\ERHYM\test_data\ERHYM_Validation_ShortFormAB_v1.dat", false);

           


            //test data
            inputfilePaths[0] = PreparePath(rootFolderforDatfiles, @"\ELNMC", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[0], "*.dat"));

            inputfilePaths[1] = PreparePath(rootFolderforDatfiles, @"\ELNFR", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[1], "*.dat"));

            inputfilePaths[2] = PreparePath(rootFolderforDatfiles, @"\ELSMC", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[2], "*.dat"));

            inputfilePaths[3] = PreparePath(rootFolderforDatfiles, @"\ELSFR", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[3], "*.dat"));

            inputfilePaths[4] = PreparePath(rootFolderforDatfiles, @"\ERHYM", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[4], "*.dat"));

            inputfilePaths[5] = PreparePath(rootFolderforDatfiles, @"\EBLMC", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[5], "*.dat"));

            inputfilePaths[6] = PreparePath(rootFolderforDatfiles, @"\EBLFR", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[6], "*.dat"));

            inputfilePaths[7] = PreparePath(rootFolderforDatfiles, @"\EVOCB", false);
            files.Add(System.IO.Directory.GetFiles(inputfilePaths[7], "*.dat"));


            for (int i = 0; i < files.Count; i++)
            {
                for (int j = 0; j < files[i].Length; j++)
                {
                    ReadExcel(files[i][j], id, timePoint);
                }

            }


            //for (int i = 0; i < 8; i++)
            //    ReadExcel(inputfilePaths[i], timePoint);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void timePoint1_CheckedChanged(object sender, EventArgs e)
        {

            if (timePoint1.Checked == true)
                timePoint = 1;
        }

        private void timePoint2_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint2.Checked == true)
                timePoint = 2;
        }

        private void timePoint3_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint3.Checked == true)
                timePoint = 3;
        }

        private void timePoint4_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint4.Checked == true)
                timePoint = 4;
        }

        private void timePoint5_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint5.Checked == true)
                timePoint = 5;
        }

        private void timePoint6_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint6.Checked == true)
                timePoint = 6;
        }

        private void timePoint7_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint7.Checked == true)
                timePoint = 7;
        }

        private void timePoint8_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint8.Checked == true)
                timePoint = 8;
        }

        private void timePoint9_CheckedChanged(object sender, EventArgs e)
        {
            if (timePoint9.Checked == true)
                timePoint = 9;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}