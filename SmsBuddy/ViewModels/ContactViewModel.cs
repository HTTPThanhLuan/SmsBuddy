using Microsoft.Win32;
using NullVoidCreations.WpfHelpers.Commands;
using SmsBuddy.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows.Input;

namespace SmsBuddy.ViewModels
{
    class ContactViewModel: ChildViewModelBase
    {
        ContactModel _contact;
        IEnumerable<ContactModel> _contacts;
        string _newMobile, _selectedMobile;
        ICommand _refresh, _new, _save, _delete, _addMobile, _removeMobile, _importFile;

        DataTable _dt;
        private String _fileImport;

        public ContactViewModel() : base("Contacts", "contacts-32.png") {
            Dt = new DataTable();
        }

        #region properties

        public string NewMobileNumber
        {
            get { return _newMobile; }
            set { Set(nameof(NewMobileNumber), ref _newMobile, value); }
        }

        public string SelectedMobileNumber
        {
            get { return _selectedMobile; }
            set { Set(nameof(SelectedMobileNumber), ref _selectedMobile, value); }
        }

        public ContactModel Contact
        {
            get { return _contact; }
            set { Set(nameof(Contact), ref _contact, value); }
        }

        public IEnumerable<ContactModel> Contacts
        {
            get { return _contacts; }
            private set { Set(nameof(Contacts), ref _contacts, value); }
        }

        public DataTable Dt
        {
            get { return _dt; }
            private set { Set(nameof(Dt), ref _dt, value); }
        }
        public String FileImport
        {
            get { return _fileImport; }
            set { Set(nameof(FileImport), ref _fileImport, value); }
        }

        #endregion

        #region commands

        public ICommand RefreshCommand
        {
            get
            {
                if (_refresh == null)
                    _refresh = new RelayCommand(Refresh);

                return _refresh;
            }
        }

        public ICommand NewCommand
        {
            get
            {
                if (_new == null)
                    _new = new RelayCommand(New);

                return _new;
            }
        }

        public ICommand SaveCommand
        {
            get
            {
                if (_save == null)
                    _save = new RelayCommand(Save);

                return _save;
            }
        }

        public ICommand DeleteCommand
        {
            get
            {
                if (_delete == null)
                    _delete = new RelayCommand(Delete);

                return _delete;
            }
        }

        public ICommand AddMobileCommand
        {
            get
            {
                if (_addMobile == null)
                    _addMobile = new RelayCommand(AddMobile) { IsSynchronous = true };

                return _addMobile;
            }
        }

        public ICommand RemoveMobileCommand
        {
            get
            {
                if (_removeMobile == null)
                    _removeMobile = new RelayCommand(RemoveMobile) { IsSynchronous = true };

                return _removeMobile;
            }
        }

        public ICommand ImportFileCommand
        {
            get
            {
                if (_importFile == null)
                    _importFile = new RelayCommand(ImportFile);

                return _importFile;
            }
        }

        #endregion

        void AddMobile()
        {
            ErrorMessage = null;

            if (Contact == null)
                ErrorMessage = "Select or create a new message first.";
            else if (string.IsNullOrEmpty(NewMobileNumber))
                ErrorMessage = "Enter mobile number.";
            else
            {
                Contact.MobileNumbers.Add(NewMobileNumber);
                NewMobileNumber = null;
            }
        }

        void RemoveMobile()
        {
            ErrorMessage = null;

            if (Contact == null)
                ErrorMessage = "Select or create a new message first.";
            else if (SelectedMobileNumber == null)
                ErrorMessage = "Select a mobile number to remove.";
            else
            {
                if (Contact.MobileNumbers.Remove(SelectedMobileNumber))
                    SelectedMobileNumber = null;
            }
        }

        void Delete()
        {
            ErrorMessage = null;

            if (Contact == null)
                ErrorMessage = "Select or create a new message first.";
            else
            {
                Contact.Delete();
                Refresh();
                New();
            }
        }

        void New()
        {
            ErrorMessage = null;
            Contact = new ContactModel();
        }

        void Save()
        {
            ErrorMessage = null;

            if (Contact == null)
                ErrorMessage = "Select or create a new contact.";
            else if (string.IsNullOrEmpty(Contact.Name))
                ErrorMessage = "Name not specified.";
            else if (Contact.MobileNumbers.Count == 0)
                ErrorMessage = "Mobile number not specified.";
            else
            {
                Contact.Save();
                Refresh();
                New();

                if(Dt.Rows.Count != 0)
                {
                    if (Dt.Rows.Count != 0)
                    {
                        foreach (DataRow dtRow in Dt.Rows)
                        {
                            
                            var mobileList = dtRow["Mobile Numbers"].ToString().Trim().Split(',');
                            var mobileCollection = new ObservableCollection<string>(mobileList);

                            Contact.FirstName = dtRow["First Name"].ToString();
                            Contact.LastName = dtRow["Last Name"].ToString();
                            Contact.Company = dtRow["Company"].ToString();
                            Contact.MobileNumbers = mobileCollection;

                            Contact.Save();
                            Refresh();
                            New();
                        }

                    }
                }
            }
        }

        void Refresh()
        {
            ErrorMessage = null;
            Contacts = new ContactModel().Get() as IEnumerable<ContactModel>;

            FileImport = String.Empty;
        }

        [STAThread]
        void ImportFile()
        {
            OpenFileDialog openFileDialogContact = new OpenFileDialog();
            openFileDialogContact.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialogContact.Filter = "Excel Files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm";

            if (openFileDialogContact.ShowDialog() == true)
            {
                //1. Đọc và ghi file vào DataTable

                var filePath = openFileDialogContact.FileName;
                //* Code này đọc được Excel Office Interop
                // http://www.codescratcher.com/wpf/import-excel-file-datagrid-wpf/
                //*-------------------------------Microsoft.Office.Interop.Excel-SETUP-START----------------------------------*//
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(filePath.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                //*-------------------------------Microsoft.Office.Interop.Excel-SETUP-END-----------------------------------*//

                Dt = new DataTable();

                var fileExtension = Path.GetExtension(openFileDialogContact.FileName);
                if (fileExtension == ".xlsx" || fileExtension == ".xls")
                {
                    FileImport += openFileDialogContact.FileName + " | ";

                    #region EPPlus
                    //using (FileStream stream = File.Open(filePath, FileMode.Open))
                    //{

                    //}

                    //* Code này không đọc được Excel Office Interop
                    //FileInfo excel = new FileInfo(filePath);

                    //var package = new ExcelPackage(excel);
                    ////ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    //ExcelWorksheet sheet = package.Workbook.Worksheets["Worksheet"];

                    //int rowcount = sheet.Dimension.Rows;
                    //int columncount = sheet.Dimension.Columns;

                    //DataTable dt = new DataTable();

                    //for (int col = 1; col <= columncount; col++)
                    //{
                    //    dt.Columns.Add(sheet.Cells[1, col].Value.ToString());
                    //}

                    //for (int i = 2; i <= rowcount; i++)
                    //{
                    //    DataRow dtr = dt.NewRow();

                    //    for (int j = 1; j <= columncount; j++)
                    //    {
                    //        if (sheet.Cells[i, j].Value == null)
                    //            continue;
                    //        else
                    //        {
                    //            dtr[j - 1] = sheet.Cells[i, j].Value.ToString();
                    //        }
                    //    }
                    //    dt.Rows.Add(dtr);
                    //}

                    #endregion

                    int rowCnt = 0;
                    int colCnt = 0;
                    string strCellData = "";
                    double douCellData;

                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        string strColumn = "";
                        strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        Dt.Columns.Add(strColumn, typeof(string));
                    }

                    for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                    {
                        DataRow dtr = Dt.NewRow();
                        for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                        {
                            try
                            {

                                strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                                dtr[colCnt - 1] = strCellData;
                            }
                            catch (Exception ex)
                            {
                                douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            }
                        }

                        Dt.Rows.Add(dtr);
                    }

                }

                excelBook.Close(true, null, null);
                excelApp.Quit();

                ErrorMessage = "Imported success!";
            }
        }
    }
}
