using Microsoft.Win32;
using NullVoidCreations.WpfHelpers.Commands;
using NullVoidCreations.WpfHelpers.DataStructures;
using OfficeOpenXml;
using SmsBuddy.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace SmsBuddy.ViewModels
{
    class SmsViewModel: ChildViewModelBase, INotifyPropertyChanged
    {
        //private static object _lock = new object();

        SmsModel _sms;
        ContactModel _selectedContact, _selectedContact1;
        IEnumerable<TemplateModel> _templates;
        IEnumerable<SmsGatewayBase> _gateways;
        IEnumerable<SmsModel> _messages;
        IEnumerable<ContactModel> _contacts;
        IEnumerable<int> _hours, _minutes;
        ICommand _refresh, _new, _send, _save, _delete, _addMobile, _removeMobile, _importFile, _clearMessage;

        DataTable _dt;
        private String _fileImport;

        public SmsViewModel() : base("Messages", "sms-32.png") {
            Sms = new SmsModel();
            Dt = new DataTable();
        }

        #region properties

        public ContactModel SelectedContact
        {
            get { return _selectedContact; }
            set { Set(nameof(SelectedContact), ref _selectedContact, value); }
        }

        public ContactModel SelectedContact1
        {
            get { return _selectedContact1; }
            set { Set(nameof(SelectedContact1), ref _selectedContact1, value); }
        }

        public SmsModel Sms
        {
            get { return _sms; }
            set { Set(nameof(Sms), ref _sms, value); }
        }

        public IEnumerable<TemplateModel> Templates
        {
            get { return _templates; }
            private set { Set(nameof(Templates), ref _templates, value); }
        }

        public IEnumerable<SmsModel> Messages
        {
            get { return _messages; }
            private set { Set(nameof(Messages), ref _messages, value); }
        }

        public IEnumerable<SmsGatewayBase> Gateways
        {
            get { return _gateways; }
            private set { Set(nameof(Gateways), ref _gateways, value); }
        }

        public IEnumerable<ContactModel> Contacts
        {
            get { return _contacts; }
            private set { Set(nameof(Contacts), ref _contacts, value); }
        }

        public IEnumerable<int> Hours
        {
            get
            {
                if (_hours == null)
                {
                    var hours = new List<int>();
                    for (var hour = 0; hour < 24; hour++)
                        hours.Add(hour);
                    _hours = hours;
                }

                return _hours;
            }
        }

        public IEnumerable<int> Minutes
        {
            get
            {
                if (_minutes == null)
                {
                    var minutes = new List<int>();
                    for (var minute = 0; minute < 60; minute++)
                        minutes.Add(minute);
                    _minutes = minutes;
                }

                return _minutes;
            }
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
                    _new = new RelayCommand<SmsModel>(New);

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

        public ICommand SendCommand
        {
            get
            {
                if (_send == null)
                    _send = new RelayCommand(Send);

                return _send;
            }
        }

        public ICommand AddMobileNumberCommand
        {
            get
            {
                if (_addMobile == null)
                    _addMobile = new RelayCommand<List<object>>(AddMobile) { IsSynchronous = true };

                return _addMobile;
            }
        }

        public ICommand RemoveMobileNumberCommand
        {
            get
            {
                if (_removeMobile == null)
                    _removeMobile = new RelayCommand<List<object>>(RemoveMobile) { IsSynchronous = true };

                return _removeMobile;
            }
        }

        public ICommand ImportFileCommand
        {
            get
            {
                if (_importFile == null)
                    _importFile = new RelayCommand<List<object>> (ImportFile) { IsSynchronous = true };

                return _importFile;
            }
        }
        public ICommand ClearMessageCommand
        {
            get
            {
                if (_clearMessage == null)
                    _clearMessage = new RelayCommand(ClearMessage);

                return _clearMessage;
            }
        }

        #endregion

        void RemoveMobile(List<object> arguments)
        {
            var number = arguments[0] as string;
            var collection = arguments[1] as ObservableCollection<string>;
            collection.Remove(number);
        }

        void AddMobile(List<object> arguments)
        {
            var contact = arguments[0] as ContactModel;
            if(arguments[1] == DependencyProperty.UnsetValue)
            {
                ErrorMessage = "Click 'New' button if you want to send new message.";
            }
            else
            {
                var collection = arguments[1] as ObservableCollection<string>;
                foreach (var number in contact.MobileNumbers)
                    if (!collection.Contains(number))
                        collection.Add(number);
            }
        }

        void Delete()
        {
            ErrorMessage = null;

            if (Sms == null)
                ErrorMessage = "Select or create a new message first.";
            else
            {
                Sms.Delete();
                Refresh();
                New(null);
            }
        }

        void ClearMessage()
        {
            ErrorMessage = null;

            if (Messages == null)
                ErrorMessage = "Select or create a new message first.";
            else
            {
                Sms.DeleteAll();
                Delete();
            }
        }

        void New(SmsModel sms = null)
        {
            var schedule = DateTime.Now.AddMinutes(15);
            ErrorMessage = null;
            if(sms == null)
            {
                Sms = new SmsModel
                {
                    RepeatDaily = true,
                    Hour = schedule.Hour,
                    Minute = schedule.Minute,
                };

                //BindingOperations.EnableCollectionSynchronization(Sms.MobileNumbers, _lock);
            }
            else
            {
                if (sms.Id != 0)
                    sms.Id += 1;
                sms.RepeatDaily = true;
                sms.Hour = schedule.Hour;
                sms.Minute = schedule.Minute;
                sms.MobileNumbers = new ExtendedObservableCollection<string>();
            }
        }

        void Save()
        {
            ErrorMessage = null;

            if (Sms == null)
                ErrorMessage = "Select or create a new message.";
            else if (Sms.MobileNumbers.Count == 0)
                ErrorMessage = "Mobile number not specified.";
            else if (Sms.Gateway == null)
                ErrorMessage = "SMS gateway not selected.";
            else if (Sms.Template == null)
                ErrorMessage = "Template not selected.";
            else if (string.IsNullOrEmpty(Sms.Message))
                ErrorMessage = "Message not specified.";
            else if (Sms.RepeatDaily && Sms.MobileNumbersScheduled.Count == 0)
                ErrorMessage = "Mobile number for scheduled SMS sending not specified.";
            else
            {
                ErrorMessage = "Saving....";

                var listMobileNumber = new List<string>(Sms.MobileNumbers);

                foreach (var number in listMobileNumber)
                {

                    App.Current.Dispatcher.Invoke((Action)delegate
                    {
                        New(Sms);
                        Sms.MobileNumbers.Add(number);
                        Sms.Save();
                        //Refresh();
                    });

                }

                if (Dt.Rows.Count != 0)
                {
                    foreach (DataRow dtRow in Dt.Rows)
                    {
                        App.Current.Dispatcher.Invoke((Action)delegate
                        {
                            New(Sms);
                            Sms.MobileNumbers.Add(dtRow["Mobile Numbers"].ToString());
                            Sms.Save();
                            //Refresh();
                        });

                    }

                }

                Refresh();
                New(null);
                ErrorMessage = "Save success!";
                //Sms.Save();
                //Refresh();
                //New();
            }
        }

        void Send()
        {
            ErrorMessage = null;

            if (Sms == null)
                ErrorMessage = "Select or create a new message.";
            else if (Sms.MobileNumbers.Count == 0)
                ErrorMessage = "Mobile number not specified.";
            else if (Sms.Gateway == null)
                ErrorMessage = "SMS gateway not selected.";
            else if (Sms.Template == null)
                ErrorMessage = "Template not selected.";
            else if (string.IsNullOrEmpty(Sms.Message))
                ErrorMessage = "Message not specified.";
            else
            {
                foreach(SmsModel sms in Messages)
                {
                    var sentMessage = sms.Gateway.Send(sms, sms.MobileNumbers);
                    ErrorMessage = sentMessage.IsSent ? null : sentMessage.GatewayMessage;
                    sentMessage.Save();
                }
                //var sentMessage = Sms.Gateway.Send(Sms, Sms.MobileNumbers);
                //ErrorMessage = sentMessage.IsSent ? null : sentMessage.GatewayMessage;
                //sentMessage.Save();
            }
        }

        void Refresh()
        {
            ErrorMessage = null;
            
            Gateways = Shared.Instance.Database.GetCollection<SmsGatewayBase>().FindAll();
            Contacts = new ContactModel().Get() as IEnumerable<ContactModel>;
            Templates = new TemplateModel().Get() as IEnumerable<TemplateModel>;
            Messages = new SmsModel().Get() as IEnumerable<SmsModel>;

            if (Sms.MobileNumbers.Count != 0)
            {
                Sms.MobileNumbers = new ExtendedObservableCollection<string>();
                //Task.Factory.StartNew(Sms.MobileNumbers.Clear);
            }

            FileImport = String.Empty;
            SelectedContact = new ContactModel();
            SelectedContact1 = new ContactModel();

        }
        [STAThread]
        void ImportFile(List<object> arguments)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = "Excel Files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm";

            var contact = arguments[0] as TemplateModel;
            var message = arguments[1] as string;          

            if (openFileDialog.ShowDialog() == true)
            {
                //1. Đọc và ghi file vào DataTable

                var filePath = openFileDialog.FileName;
                //* Code này đọc được Excel Office Interop
                // http://www.codescratcher.com/wpf/import-excel-file-datagrid-wpf/
                //*-------------------------------Microsoft.Office.Interop.Excel-SETUP-START----------------------------------*//
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(filePath.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                //*-------------------------------Microsoft.Office.Interop.Excel-SETUP-END-----------------------------------*//

                Dt = new DataTable();

                var fileExtension = Path.GetExtension(openFileDialog.FileName);
                if (fileExtension == ".xlsx" || fileExtension == ".xls")
                {
                    FileImport += openFileDialog.FileName + " | ";

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
