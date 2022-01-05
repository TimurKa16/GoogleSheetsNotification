// GoogleSheets 8

using System;
using Microsoft.Win32;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;

using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace NotificationMarat
{



    public partial class MainFormMarat : Form
    {
        //private static string ClientSecret = "C:\\Users\\USER\\source\\repos\\GoogleSheetsApi\\Test\\GoogleSheets.json";
        private static string ClientSecret = "C:\\Program Files (x86)\\Timur Corporation\\TimurNotification\\GoogleSheets.json";
        private static readonly string[] ScopesSheets = { SheetsService.Scope.Spreadsheets };
        private static readonly string AppName = "";
        private static readonly string SpreadsheetId = "";
        private const string Range = "'Узбекистан'!A1:10000";
        private const int SheetId = 78949560;



        const int paymentDateColumn = 23 - 1;         // Колонка с датами
        const int paymentSignColumn = 22 - 1;  // Колонка с плюсиком
        const int driverColumn = 16 - 1;
        int customerColumn = 11 - 1;

        string paymentSign = "0,00";
        int paymentSignLength = 6;
        int minLengthOfRow = 16;


        private static string[,] Data =
        {
            {"AA", "AB"},
            {"Ba", "BB"}
        };
        


        //string[][] result;
        //string[][] oldResult;
        List<string[]> filteredResults = new List<string[]>();
        public List<List<Note>> notes = new List<List<Note>>();
        public List<List<Note>> oldNotes = new List<List<Note>>();
        public List<Note> redNotes = new List<Note>();

        List<string[][]> sheetResults = new List<string[][]>();
        List<string[][]> oldSheetResults = new List<string[][]>();


        bool debugMode = false;

        Thread myThread;

        List<Request> requests = new List<Request>();
        UserCredential credential;
        SheetsService service;


        bool[] overallIsChanged;
        bool[] localChange;
        bool[] writingIsAllowed;
        bool[] previousIsChanged; // means that next reading we are expecting cells are not changed
        bool[] nowIsChanged;
        

        public MainFormMarat()
        {
            InitializeComponent();
            //Visible = false;
            //Hide();



            Text = string.Empty;
            ControlBox = false;

            if (debugMode)
                SetAutorunValue(false);
            else
                SetAutorunValue(true);

            SetSpreadsheets();

            for (int i = 0; i < sheets.Count; i++)
            {
                notes.Add(new List<Note>());
                oldNotes.Add(new List<Note>());
                oldSheetResults.Add(new string[1][]);
            }


            overallIsChanged = new bool[sheets.Count];
            localChange = new bool[sheets.Count];
            writingIsAllowed = new bool[sheets.Count];
            previousIsChanged = new bool[sheets.Count]; // means that next reading we are expecting cells are not changed
            nowIsChanged = new bool[sheets.Count];

            string s = null;
            bool redNotesExist = false;

                    HandleGoogleSheets();
                    
                    HandleUserTable();

        }




        private void MainFormLoad(object sender, EventArgs e)
        {

        }

        DateTime CorrectDate(string date)
        {
            string[] buf = date.Split(' ', '\\', '/', '.', ',', '_');
            if (buf.Length < 3)
                return (DateTime.MinValue);
            int day = Convert.ToInt16(buf[0]);
            int month = Convert.ToInt16(buf[1]);
            int year = Convert.ToInt16(buf[2]);
            if (year < 2000)
                year += 2000;

            try
            {
                return (new DateTime(year, month, day));
            }
            catch (Exception)
            {
                return (DateTime.MinValue);
            };
        }

        bool NotesAreEqual(List<Note> notes1, List<Note> notes2, string[][] spreadsheet1, string[][] spreadsheet2)
        {
            int j = 0;
            if (spreadsheet2 == null)
                return false;
            else if (spreadsheet2[0] == null)
                return false;

            if (notes1.Count != notes2.Count)
                return false;
            else
            {
                    for (int i = 0; i < notes1.Count; i++)
                    {
                        if (spreadsheet1[notes1[i].rowNumber].Length >= paymentSignColumn + 1)
                        {
                            if (spreadsheet1[notes1[i].rowNumber].Length >= driverColumn && spreadsheet2[notes1[i].rowNumber].Length >= driverColumn)
                                if (spreadsheet1[notes1[i].rowNumber][driverColumn] != spreadsheet2[notes1[i].rowNumber][driverColumn])
                                    return false;
                            if (spreadsheet1[notes1[i].rowNumber].Length >= paymentSignColumn && spreadsheet2[notes1[i].rowNumber].Length >= paymentSignColumn)
                                if (spreadsheet1[notes1[i].rowNumber][paymentSignColumn] != spreadsheet2[notes1[i].rowNumber][paymentSignColumn])
                                    return false;

                            if (spreadsheet1[notes1[i].rowNumber].Length >= paymentDateColumn + 1)
                            {
                                if (spreadsheet1[notes1[i].rowNumber].Length >= paymentDateColumn && spreadsheet2[notes1[i].rowNumber].Length >= paymentDateColumn)
                                    if (spreadsheet1[notes1[i].rowNumber][paymentDateColumn] != spreadsheet2[notes1[i].rowNumber][paymentDateColumn])
                                        return false;
                            }
                        }
                    }
            }
            return true;
        }

        static int credentialError = 0;

        private void HandleGoogleSheets()
        {
            bool readingIsOk = false;
            while (!readingIsOk)
            {
                try
                {
                    // Получаем доступ

                    credential = GetSheetCredentials();
                    service = GetService(credential);
                    readingIsOk = true;
                }
                catch (Exception)
                {
                    credentialError++;
                    if (credentialError > 4)
                    {
                        MessageBox.Show("Программа не установлена");

                        if (myThread != null)
                            myThread.Abort();


                        Process.GetCurrentProcess().Kill();
                    }
                    Thread.Sleep(1000);

                }
            }


            // Insert here
            for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++)
            {

                //try
                {
                    readingIsOk = false;

                    while (!readingIsOk)
                    {
                        try
                        {
                            // Считываем строки

                            sheetResults.Add(GetCells(service, sheets[sheetIndex].range,
                                sheets[sheetIndex].spreadsheetId));
                            readingIsOk = true;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(2000);
                        };
                    }




                    if (sheetIndex == 1)
                    {
                        sheetIndex++;
                        sheetIndex--;
                    }

                    // Фильтруем записи
                    FilterNotes(sheets[sheetIndex], notes[sheetIndex],
                        sheetResults[sheetIndex]);

                    if (sheetIndex == 1)
                    {
                        sheetIndex++;
                        sheetIndex--;
                    }

                    bool notesAreEqual = true;
                    if (sheetResults.Count > sheetIndex)
                    {
                        if (oldSheetResults.Count > sheetIndex)
                        {
                            notesAreEqual = NotesAreEqual(notes[sheetIndex], oldNotes[sheetIndex],
                                sheetResults[sheetIndex], oldSheetResults[sheetIndex]);
                        }
                        else
                        {
                            notesAreEqual = false;
                        }
                    }

                    if (!notesAreEqual)
                    {
                        if (sheetIndex == 1)
                        {
                            sheetIndex++;
                            sheetIndex--;
                        }
                        nowIsChanged[sheetIndex] = true;
                        writingIsAllowed[sheetIndex] = false;
                    }
                    else
                    {
                        nowIsChanged[sheetIndex] = false;

                        if (previousIsChanged[sheetIndex])
                            writingIsAllowed[sheetIndex] = true;
                        else
                            writingIsAllowed[sheetIndex] = false;
                    }

                    previousIsChanged[sheetIndex] = nowIsChanged[sheetIndex];

                    if (writingIsAllowed[sheetIndex])
                    {
                        WriteToSpreadsheet(notes[sheetIndex], sheets[sheetIndex]);
                        writingIsAllowed[sheetIndex] = false;
                    }

                    if (requests.Count != 0)
                    {
                        try
                        {
                            //List<Request> tmpRequests = new List<Request>();
                            // for (int i = 0; i < requests.Count; i++)
                            {
                                //tmpRequests.Add(requests[i]);

                                //if (i == 200 - 1)
                                {
                                    BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
                                    batchUpdateRequest.Requests = requests;
                                    service.Spreadsheets.BatchUpdate(batchUpdateRequest, SpreadsheetId).Execute();


                                    //tmpRequests = new List<Request>();

                                    //Thread.Sleep(2000);
                                }
                                requests = new List<Request>();
                            }
                        }
                        catch (Exception) { MessageBox.Show(""); }
                    }

                    oldNotes = new List<List<Note>>(notes);

                    CopyMatrixes(ref oldSheetResults, sheetIndex, sheetResults[sheetIndex]);
                }

            }
        }

        void CopyMatrixes(ref List<string[][]> list, int index, string[][] matrix2)
        {
            list[index] = new string[matrix2.Length][];

            for (int i = 0; i < list[index].Length; i++)
                list[index] = matrix2;
        }

        private void HandleUserTable()
        {
            redNotes = new List<Note>();

            for (int googleSpreadsheetIndex = 0; googleSpreadsheetIndex < notes.Count; googleSpreadsheetIndex++)
            {
                for (int i = 0; i < notes[googleSpreadsheetIndex].Count; i++)
                    if (notes[googleSpreadsheetIndex][i].status == NoteStatus.Red)
                        redNotes.Add(notes[googleSpreadsheetIndex][i]);
            }

            dataGridView1.DataSource = redNotes;

            dataGridView1.Columns[0].HeaderText = "Строка";
            dataGridView1.Columns[0].Visible = false;

            dataGridView1.Columns[1].HeaderText = "Заказчик";

            dataGridView1.Columns[2].HeaderText = "Страна";

            dataGridView1.Columns[3].HeaderText = "Водитель";

            dataGridView1.Columns[4].HeaderText = "Дата оплаты";

            dataGridView1.Columns[5].HeaderText = "Статус";
            dataGridView1.Columns[5].Visible = false;

            dataGridView1.Columns[6].HeaderText = "Сумма";
            dataGridView1.Columns[6].Visible = false;

        }




        private static UserCredential GetSheetCredentials()
        {
            using (var stream = new FileStream(ClientSecret, FileMode.Open, FileAccess.Read))
            {
                var credPath = Path.Combine(Directory.GetCurrentDirectory(), "C:\\Users\\Public\\Documents\\GoogleSheets");

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    ScopesSheets,
                    "User",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
        }


        private static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName
            });
        }

        private static void FillSpreadsheet(SheetsService service, string spreadsheetId, int? sheetId, Note note, int paymentDateColumn, string s, ref List<Request> requests)
        {
            List<CellData> values = new List<CellData>();



            values.Add(new CellData
            {
                UserEnteredValue = new ExtendedValue
                {
                    StringValue = s
                }
            });

            //create request with
            requests.Add(
                new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Start = new GridCoordinate
                        {
                            SheetId = sheetId,
                            RowIndex = note.rowNumber,
                            ColumnIndex = paymentDateColumn
                        },
                        Rows = new List<RowData> { new RowData { Values = values } },
                        Fields = "userEnteredValue"
                    }
                }
                );
        }



        private static string[][] GetCells(SheetsService service, string range, string spreadSheetId)
        {
            SpreadsheetsResource.ValuesResource.GetRequest getRequest = service.Spreadsheets.Values.Get(spreadSheetId, range);

            ValueRange response = null;

            bool readingIsOk = false;
            IAsyncResult res;
            while (!readingIsOk)
            {
                Action action = () =>
                {
                    try
                    {
                        response = getRequest.Execute();
                    }
                    catch(Exception)
                    {
                        Thread.Sleep(1000);
                        readingIsOk = false;
                    }
                };

                res = action.BeginInvoke(null, null);

                if (res.AsyncWaitHandle.WaitOne(3000))
                    readingIsOk = true;
                else
                {
                    Thread.Sleep(1000);
                    readingIsOk = false;
                }

            }

            

            if (response.Values != null)
            {
                string[][] result = new string[response.Values.Count][];
                for (int i = 0; i < response.Values.Count; i++)
                    result[i] = new string[response.Values[i].Count];



                if (response.Values != null)
                {
                    for (int i = 0; i < response.Values.Count; i++)
                    {
                        if (response.Values[i] != null)
                        {
                            for (int j = 0; j < response.Values[i].Count; j++)
                            {
                                result[i][j] = response.Values[i][j].ToString();
                            }

                        }
                    }

                }


                return result;
            }
            else
                return null;
        }

        private static void SetFormatForElements(SheetsService service, string spreadsheetId, int? sheetId, int row, int column, NoteStatus status, ref List<Request> requests)
        {

            List<CellData> values = new List<CellData>
            {
                new CellData
                {
                    UserEnteredFormat = GetCellFormat(status)
                }
            };

            var updateCellRequest = new Request
            {
                UpdateCells = new UpdateCellsRequest
                {

                    Start = new GridCoordinate
                    {

                        SheetId = sheetId,
                        RowIndex = row,
                        ColumnIndex = column
                    },
                    Rows = new List<RowData> { new RowData { Values = values } },
                    Fields = "userEnteredFormat"
                }
            };


            requests.Add(updateCellRequest);

        }

        private static CellFormat GetCellFormat(NoteStatus status)
        {
            CellFormat cellFormat = new CellFormat();
            cellFormat.Borders = new Borders
            {
                Bottom = new Border
                {
                    Style = "SOLID",
                    Color = new Color
                    {
                        Red = 0.1f,
                        Green = 0.1f,
                        Blue = 0.1f
                    }
                },

                Top = new Border
                {
                    Style = "SOLID",
                    Color = new Color
                    {
                        Red = 0.1f,
                        Green = 0.1f,
                        Blue = 0.1f
                    }
                },

                Left = new Border
                {
                    Style = "SOLID",
                    Color = new Color
                    {
                        Red = 0.1f,
                        Green = 0.1f,
                        Blue = 0.1f
                    }
                },

                Right = new Border
                {
                    Style = "SOLID",
                    Color = new Color
                    {
                        Red = 0.1f,
                        Green = 0.1f,
                        Blue = 0.1f
                    }
                },
            };

            cellFormat.TextFormat = new TextFormat
            {
                ForegroundColor = new Color
                {
                    Red = 0.1f,
                    Green = 0.1f,
                    Blue = 0.1f
                }
            };

            cellFormat.HorizontalAlignment = "CENTER";
            cellFormat.VerticalAlignment = "MIDDLE";

            //cellFormat.NumberFormat

            if (status == NoteStatus.White)
            {
                cellFormat.BackgroundColor = new Color
                {
                    Red = 1f,
                    Green = 1f,
                    Blue = 1f
                };
            }
            if (status == NoteStatus.Green)
            {
                cellFormat.BackgroundColor = new Color
                {
                    Red = 144f,
                    Green = 83f,
                    Blue = 185f
                };
            }
            else if (status == NoteStatus.Yellow)
            {
                cellFormat.BackgroundColor = new Color
                {
                    Red = 1f,
                    Green = 1f,
                    Blue = 0f
                };
            }
            else if (status == NoteStatus.Red)
            {
                cellFormat.BackgroundColor = new Color
                {
                    Red = 1f,
                    Green = 0f,
                    Blue = 0f
                };
            }

            return cellFormat;
        }




        void FilterNotes(GoogleSpreadsheet googleSpreadsheet, List<Note> notes, string[][] result)
        {
            bool isPaid = false;

            for (int i = 1; i < result.Length; i++)
            {
                if (i == 118)
                {
                    i++;
                    i--;
                }
                try
                {
                    if (result[i].Length >= minLengthOfRow)
                    {
                        if (result[i].Length > paymentSignColumn)
                        {
                            int tmpLength = result[i][paymentSignColumn].Length;

                            if (
                                ((result[i][paymentSignColumn].Contains(paymentSign)) &&
                                result[i][paymentSignColumn].Length == paymentSignLength) ||

                                result[i][paymentSignColumn] == "" ||
                                result[i][paymentSignColumn] == "0" ||
                                result[i][paymentSignColumn] == "  -   ₽ "
                                )
                                isPaid = false;
                            else
                                isPaid = true;

                            DateTime resultTime;

                            //if (result[i].Length >= paymentDateColumn)
                            {
                                if (isPaid == true)
                                {
                                    notes.Add(new Note());
                                    notes[notes.Count - 1].rowNumber = i;
                                    notes[notes.Count - 1].status = NoteStatus.Green;
                                    notes[notes.Count - 1].driver = result[i][driverColumn];
                                    notes[notes.Count - 1].paymentSign = result[i][paymentSignColumn];
                                    notes[notes.Count - 1].country = googleSpreadsheet.name;
                                    notes[notes.Count - 1].customer = result[i][customerColumn];
                                    try
                                    {
                                        resultTime = CorrectDate(result[i][paymentDateColumn]);
                                        notes[notes.Count - 1].date = resultTime;
                                    }
                                    catch (Exception)
                                    {
                                        notes[notes.Count - 1].date = DateTime.MinValue;
                                    };
                                }
                                else
                                {

                                    notes.Add(new Note());
                                    notes[notes.Count - 1].rowNumber = i;
                                    try
                                    {
                                        resultTime = CorrectDate(result[i][paymentDateColumn]);
                                        notes[notes.Count - 1].paymentSign = result[i][paymentSignColumn];
                                        notes[notes.Count - 1].country = googleSpreadsheet.name;
                                        notes[notes.Count - 1].customer = result[i][customerColumn];

                                        if (resultTime.Subtract(DateTime.Now).TotalDays <= 1)
                                        {
                                            notes[notes.Count - 1].date = resultTime;

                                            //@@@ Driver
                                            //@@@ Organization
                                            if (resultTime.Subtract(DateTime.Now).TotalDays <= 0)
                                                notes[notes.Count - 1].status = NoteStatus.Red;
                                            else
                                                notes[notes.Count - 1].status = NoteStatus.Yellow;

                                            notes[notes.Count - 1].driver = result[i][driverColumn];
                                        }
                                        if (resultTime == DateTime.MinValue)
                                            notes[notes.Count - 1].status = NoteStatus.White;
                                    }
                                    catch (Exception)
                                    {
                                        notes[notes.Count - 1].status = NoteStatus.White;
                                    };
                                }

                            }

                        }
                        else
                        {
                            notes.Add(new Note());
                            notes[notes.Count - 1].rowNumber = i;
                            notes[notes.Count - 1].date = DateTime.MinValue;
                            notes[notes.Count - 1].driver = "";
                            notes[notes.Count - 1].status = NoteStatus.White;
                            notes[notes.Count - 1].country = googleSpreadsheet.name;
                            notes[notes.Count - 1].customer = result[i][customerColumn];
                        }
                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show(i.ToString());
                }
            }
        }

        const string applicationName = "NotificationMarat";
        public bool SetAutorunValue(bool autorun)
        {
            string ExePath = Application.ExecutablePath;
            RegistryKey reg;

            reg = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run\\");
            try
            {
                if (autorun)
                    reg.SetValue(applicationName, ExePath);
                else if (reg.ValueCount != 0)
                    reg.DeleteValue(applicationName);

                reg.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }

        void WriteToSpreadsheet(List<Note> notes, GoogleSpreadsheet googleSpreadsheet)
        {
            // Раскрашиваем ячейки
            for (int i = 0; i < notes.Count; i++)
                SetFormatForElements(service, googleSpreadsheet.spreadsheetId, googleSpreadsheet.sheetId, notes[i].rowNumber, paymentDateColumn,
                    notes[i].status, ref requests);

            string noteDate = null;
            for (int i = 0; i < notes.Count; i++)
            {
                if (notes[i].date != DateTime.MinValue)
                    noteDate = notes[i].date.ToString().Substring(0, 10);
                else
                    noteDate = "";
                FillSpreadsheet(service, googleSpreadsheet.spreadsheetId, googleSpreadsheet.sheetId, notes[i], paymentDateColumn, noteDate, ref requests);
            }
        }

        private void NewThread()
        {


            // Считываем строки
            while (true)
            {
                Thread.Sleep(5000);


                notes = new List<List<Note>>();
                sheetResults = new List<string[][]>();

                for (int i = 0; i < sheets.Count; i++)
                {
                    notes.Add(new List<Note>());
                    oldNotes.Add(new List<Note>());
                }

                HandleGoogleSheets();

            }

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (redNotes.Count == 0)
                Hide();


            //Hide();
            //MessageBox.Show("Есть 3 неоплаченные заявки");

            // создаем новый поток
            myThread = new Thread(new ThreadStart(NewThread));
            myThread.Start(); // запускаем поток
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void label2_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }



        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private List<GoogleSpreadsheet> sheets = new List<GoogleSpreadsheet>();

        void SetSpreadsheets()
        {
            sheets.Add(new GoogleSpreadsheet
            {
                name = "Туркменистан",
                spreadsheetId = SpreadsheetId,
                range = "'Туркменистан'!A1:10000",
                sheetId = 1003552189
            });

            sheets.Add(new GoogleSpreadsheet
            {
                name = "Монголия",
                spreadsheetId = SpreadsheetId,
                range = "'Монголия'!A1:10000",
                sheetId = 587930582
            });
            sheets.Add(new GoogleSpreadsheet
            {
                name = "Узбекистан",
                spreadsheetId = SpreadsheetId,
                range = "'Узбекистан'!A1:10000",
                sheetId = 78949560
            });
            sheets.Add(new GoogleSpreadsheet
            {
                name = "Казахстан",
                spreadsheetId = SpreadsheetId,
                range = "'Казахстан'!A1:10000",
                sheetId = 950179148
            });
        }

        private void dataGridView1_CellContentClick_2(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
            //Hide();

        }
    }
}
