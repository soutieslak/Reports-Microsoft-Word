using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using Novacode;

namespace Module_Report
{
    /// <summary>
    /// Externally developed class to bind spesified controls to specified functionality
    /// </summary>
    class ControlEvents_Report
    {
        #region class constructor / destructor

        Reporting_Word Report = new Reporting_Word();

        DocX document;

        /// <summary>
        /// The list that will be used to insert data into the document
        /// </summary>
        public List<Reporting_Word.Bookmarks> FinalList = new List<Reporting_Word.Bookmarks>();

        /// <summary>
        /// The list used to temporily store the unsorted data in before it goes into the final list
        /// </summary>
        public List<Reporting_Word.Bookmarks> TempList = new List<Reporting_Word.Bookmarks>();


        /// <summary>
        /// Default constructor for anything required
        /// </summary>
        public ControlEvents_Report() 
        {
            //subscribing to the event used to globaly send bitmaps to the document from the PatternDrawer Class
            //RFTools.AntennaPatterns.PatternDrawer.SendPattern += Update_ReceivedPattern;
            //subscribing to the event used to globaly send bitmaps to the document from the sweep classs
            //Overlays.Overlay.SendSweep += Update_ReceivedSweep;
            //subscribe to the event used globaly to send sweep data from the signal generator class
            //ControlEvents_SignalGenerator.SendSweepData += Update_ReceivedSweepData;
            document = Report.Load_Template("BookmarkTemplate.docx");
            FinalList = Report.List_Bookmarks(document);
        }

        /// <summary>
        /// dispose functionality of all controls when required
        /// </summary>
        public void DisposeAll()
        {
            lstMapping.Dispose();
            lstBrowser.Dispose();
            clmBookmark.Dispose();
            clmMapping.Dispose();
            clmFigure.Dispose();
            btnLoad.Dispose();
        }

        #endregion

        #region creation of all required GUI controls

        private System.Windows.Forms.ListView lstMapping;
        private System.Windows.Forms.ColumnHeader clmBookmark;
        private System.Windows.Forms.ColumnHeader clmMapping;
        private System.Windows.Forms.ListView lstBrowser;
        private System.Windows.Forms.ColumnHeader clmFigure;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnGenerate;

        #endregion

        #region definition of all binding functions

        /// <summary>
        /// Bind a listview to custom events
        /// </summary>
        /// <param name="listview"></param>
        public void bindBrowser(System.Windows.Forms.ListView listview)
        {
            lstBrowser = listview;
            lstBrowser.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(lstBrowser_ItemDrag);
            lstBrowser.MouseUp += new System.Windows.Forms.MouseEventHandler(lstBrowser_MouseUp);
        }

        /// <summary>
        /// Bind a listview to custom events
        /// </summary>
        /// <param name="listview"></param>
        public void bindMapping(System.Windows.Forms.ListView listview)
        {
            lstMapping = listview;
            lstMapping.DragDrop += new System.Windows.Forms.DragEventHandler(lstMapping_DragDrop);
            lstMapping.DragEnter += new System.Windows.Forms.DragEventHandler(lstMapping_DragEnter);
            lstMapping.MouseUp += new System.Windows.Forms.MouseEventHandler(lstMapping_MouseUp);
            lstMapping_Populate(FinalList);
        }

        /// <summary>
        /// Bind a columnheader to custom events
        /// </summary>
        /// <param name="columnheader"></param>
        public void bind_clmBookmark(System.Windows.Forms.ColumnHeader columnheader)
        {
            clmBookmark = columnheader;
        }

        /// <summary>
        /// Bind a columnheader to custom events
        /// </summary>
        /// <param name="columnheader"></param>
        public void bind_clmMapping(System.Windows.Forms.ColumnHeader columnheader)
        {
            clmMapping = columnheader;
        }

        /// <summary>
        /// Bind a columnheader to custom events
        /// </summary>
        /// <param name="columnheader"></param>
        public void bind_clmFigure(System.Windows.Forms.ColumnHeader columnheader)
        {
            clmFigure = columnheader;
        }

        /// <summary>
        /// Bind a Button to custom events
        /// </summary>
        /// <param name="button"></param>
        public void bindLoad(System.Windows.Forms.Button button)
        {
            btnLoad = button;
            btnLoad.Click += new EventHandler(btnLoad_Click);
        }

        /// <summary>
        /// Bind a Button to custom events
        /// </summary>
        /// <param name="button"></param>
        public void bindGenerate(System.Windows.Forms.Button button)
        {
            btnGenerate = button;
            btnGenerate.Click += new EventHandler(btnGenerate_Click);
        }

        #endregion

        #region All control event handlers

        /// <summary>
        /// btnLoad event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoad_Click(object sender, EventArgs e)
        {
            document = Report.Load_Template();
            //unmap existing matches
            for (int i = 0; i < FinalList.Count; i++)
                UnMap(i);
            FinalList = Report.List_Bookmarks(document);
            lstMapping_Populate(FinalList);
        }

        /// <summary>
        /// btnGenerate_Click event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            string FilePath = SaveDialog();//Gets the save location
            if (FilePath != "")
            {
                //Replace the picture bookmarks with pictures
                InsertPictures();
                //Replace the text bookmarks with data
                InsertText();
                //Create a table of context
                document.Bookmarks["bmkTableOfContext"].SetText("");
                document.InsertTableOfContents(document.Bookmarks["bmkTableOfContext"].Paragraph, "Table of context", TableOfContentsSwitches.H, "Normal", 2, null);
                //Save the report
                Report.SaveDoc(document, FilePath);
            }
            
        }

        /// <summary>
        /// Method to match and insert text next to there respective bookmarks in the document
        /// </summary>
        private void InsertText()
        {
            document.Bookmarks["bmkDate"].SetText(System.DateTime.Now.ToString());
            foreach (var bookmark in FinalList)
            {
                if (bookmark.Text != "")
                {
                    document.Bookmarks[bookmark.Name].SetText(bookmark.Text);
                }
            }
        }

        /// <summary>
        /// Method to match and insert pictures/figures to there respective bookmarks in the document
        /// </summary>
        private void InsertPictures()
        {
            foreach (var bookmark in FinalList)
            {
                if (bookmark.Figure != null)
                {
                    Picture picture;
                    if (bookmark.Name.Contains("Pattern"))
                    {
                        picture = Report.CreatePicture(ref document, bookmark.Figure, 400, 450);
                        document.Bookmarks[bookmark.Name].SetText("");
                        Console.WriteLine(document.Bookmarks[bookmark.Name].Paragraph.Text);
                        document.Bookmarks[bookmark.Name].Paragraph.AppendLine("").AppendPicture(picture);
                    }
                    if (bookmark.Name.Contains("Sweep"))
                    {
                        picture = Report.CreatePicture(ref document, bookmark.Figure, 300, 600);
                        document.Bookmarks[bookmark.Name].SetText("");
                        document.Bookmarks[bookmark.Name].Paragraph.AppendLine("").AppendPicture(picture);
                    }
                }
            }
        }

        /// <summary>
        /// Populates the list view with all available bookmarks
        /// </summary>
        /// <param name="BookmarkList">List of all the bookmarks</param>
        private void lstMapping_Populate(List<Reporting_Word.Bookmarks> BookmarkList)
        {
            lstMapping.Items.Clear();

            //Add empty subitems for each entry
            for (int i = 0; i < BookmarkList.Count; i++)
            {
                lstMapping.Items.Add(BookmarkList[i].Name);
                lstMapping.Items[i].SubItems.Add(BookmarkList[i].FilterTag);
            }
                
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstMapping_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Copy;
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstMapping_DragDrop(object sender, DragEventArgs e)
        {
            ListViewItem item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
            if (item != null)
            {
                Point pt = this.lstMapping.PointToClient(new Point(e.X, e.Y));
                ListViewItem hoveritem = this.lstMapping.GetItemAt(pt.X, pt.Y);
                if ((hoveritem != null) && (hoveritem.SubItems[1].Text == ""))
                {
                    int UnmappedIndex = item.Index;
                    int MappedIndex = hoveritem.Index;

                    Map(MappedIndex, UnmappedIndex);

                    //hoveritem.SubItems[1].Text = item.Text;
                    //lstBrowser.Items[lstBrowser.SelectedIndices[0]].BackColor = Color.LightGray;
                }
            }

        }

        /// <summary>
        /// Links an unmapped item to a mapped item
        /// </summary>
        /// <param name="MappedPos">The position if the Mapped item</param>
        /// <param name="UnmappedPos">The position of the unmapped item</param>
        private void Map(int MappedPos, int UnmappedPos)
        {
            Reporting_Word.Bookmarks FinalBookmark;
            FinalBookmark.Name = FinalList[MappedPos].Name;
            FinalBookmark.Figure = TempList[UnmappedPos].Figure;
            FinalBookmark.FilterTag = TempList[UnmappedPos].FilterTag;
            FinalBookmark.Text = TempList[UnmappedPos].Text;
            FinalBookmark.Mapped = true;

            FinalList.RemoveAt(MappedPos);
            FinalList.Insert(MappedPos, FinalBookmark);

            TempList.RemoveAt(UnmappedPos);

            //Refresh the list dispaly
            lstMapping_Populate(FinalList);
            lstBrowser_Populate(TempList);

        }


        /// <summary>
        /// Unlinks an mapped item to an unmapped
        /// </summary>
        /// <param name="MappedPos">The position if the Mapped item</param>
        private void UnMap(int MappedPos)
        {
            if (lstMapping.Items[MappedPos].SubItems[1].Text != "")
            {
                //Also tell the temp list that it has been mapped
                Reporting_Word.Bookmarks TempBookmark;
                TempBookmark.Name = "";
                TempBookmark.Figure = FinalList[MappedPos].Figure;
                TempBookmark.FilterTag = FinalList[MappedPos].FilterTag;
                TempBookmark.Text = FinalList[MappedPos].Text;
                TempBookmark.Mapped = false;

                TempList.Add(TempBookmark);

                Reporting_Word.Bookmarks FinalBookmark;
                FinalBookmark.Name = FinalList[MappedPos].Name;
                FinalBookmark.Figure = null;
                FinalBookmark.FilterTag = "";
                FinalBookmark.Text = "";
                FinalBookmark.Mapped = false;

                FinalList.RemoveAt(MappedPos);
                FinalList.Insert(MappedPos, FinalBookmark);

                //Refresh the list dispaly
                lstMapping_Populate(FinalList);
                lstBrowser_Populate(TempList);
            }
        }

        private void lstBrowser_ItemDrag(object sender, ItemDragEventArgs e)
        {
           lstBrowser.DoDragDrop(lstBrowser.SelectedItems[0], DragDropEffects.All);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstMapping_MouseUp(object sender, MouseEventArgs e)
        {
            ContextMenu Popup = new ContextMenu();
            Popup.MenuItems.Add("Unmap");
            Popup.MenuItems[0].Name = "mnuUnmap";
            Popup.MenuItems[0].Click += new EventHandler(mnuUnmap_Click);
            Popup.MenuItems[0].Enabled = false;

            if (lstMapping.Items.Count > 0)
                Popup.MenuItems[0].Enabled = true;

            if (e.Button == MouseButtons.Right)
                Popup.Show(lstMapping, new Point(e.X, e.Y));

        }

        /// <summary>
        /// Menu item used to unmap/link an item
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mnuUnmap_Click(object sender, EventArgs e)
        {
            try
            {
                int ClickedItem = lstMapping.SelectedIndices[0];
                UnMap(ClickedItem);
            }
            catch(Exception E)
            {
                MessageBox.Show("Please don't click on empty rows: \n\n" + E.ToString());
            }
        }

        private void lstBrowser_MouseUp(object sender, MouseEventArgs e)
        {
            ContextMenu Popup = new ContextMenu();
            Popup.MenuItems.Add("Add Figure from file...");
            Popup.MenuItems[0].Name = "mnuAddFigure";
            Popup.MenuItems[0].Click += new EventHandler(mnuAddFigure_Click);
            Popup.MenuItems[0].Enabled = true;
            //Remove menu item
            Popup.MenuItems.Add("Remove from list");
            Popup.MenuItems[1].Name = "mnuRemove";
            Popup.MenuItems[1].Click += new EventHandler(mnuRemove_Click);
            Popup.MenuItems[1].Enabled = true;

            if (e.Button == MouseButtons.Right)
                Popup.Show(lstBrowser, new Point(e.X, e.Y));

        }

        /// <summary>
        /// Menu item used to add images/figures from a file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mnuAddFigure_Click(object sender, EventArgs e)
        {
            //Open Dialog to let the user brows and add image files
            string[] Files = LoadFigures();
            //bmp from file
            foreach (string File in Files)
            {
                Bitmap bmp = (Bitmap)System.Drawing.Image.FromFile(File);
                Reporting_Word.Bookmarks bookmark;
                bookmark.FilterTag = System.IO.Path.GetFileName(File);//The file name
                bookmark.Figure = bmp;
                bookmark.Text = "";
                bookmark.Name = "";
                bookmark.Mapped = false;
                TempList.Add(bookmark);
            }
            lstBrowser_Populate(TempList);
        }

        /// <summary>
        /// Menu item used to remove images/figures from the unmapped list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mnuRemove_Click(object sender, EventArgs e)
        {
            //Removes the selected item from the list
            try
            {
                int ClickedItem = lstBrowser.SelectedIndices[0];
                //Remove(ClickedItem);
                TempList.RemoveAt(ClickedItem);
                lstBrowser_Populate(TempList);
            }
            catch (Exception E)
            {
                MessageBox.Show("Please don't click on empty rows: \n\n" + E.ToString());
            }

        }

        /// <summary>
        /// Opens a dialog box and allows the user to select multiple files for import
        /// </summary>
        /// <returns>Returns a string[] array containing all the file paths selected for import</returns>
        public string[] LoadFigures()
        {
            OpenFileDialog LoadFigure = new OpenFileDialog();
            LoadFigure.Title = "Load Figures for the report";
            LoadFigure.Filter = "JPG |*.jpg|PNG |*.png|All Supported Files |*.jpg;*.png|All files|*.*";
            LoadFigure.FilterIndex = 3;    //By default the All Supported Files filter is selected
            LoadFigure.Multiselect = true;

            if (LoadFigure.ShowDialog() == DialogResult.OK)
            {
                return LoadFigure.FileNames;
            }
            else return null;
        }

        /// <summary>
        /// The method that handles events triggered by the Global event 'SendPattern'
        /// </summary>
        /// <param name="FilterName">Give the data a name so that it can be filtered</param>
        /// <param name="bmp">Bitmap to be published</param>
        /// <param name="Text">Additional info</param>
        private void Update_ReceivedPattern(string FilterName, Bitmap bmp, string Text)
        {
            Reporting_Word.Bookmarks bookmark;
            bookmark.FilterTag = FilterName;
            bookmark.Figure = bmp;
            bookmark.Text = Text;
            bookmark.Name = "";
            bookmark.Mapped = false;
            TempList.Add(bookmark);
            Check_Preset();
            lstBrowser_Populate(TempList);
        }

        /// <summary>
        /// The method that handles events triggered by the Global event 'SendSweep'
        /// </summary>
        /// <param name="FilterName">Give the data a name so that it can be filtered</param>
        /// <param name="bmp">Bitmap to be published</param>
        /// <param name="Text">Additional info</param>
        private void Update_ReceivedSweep(string FilterName, Bitmap bmp, string Text)
        {
            Reporting_Word.Bookmarks bookmark;
            bookmark.FilterTag = FilterName;
            bookmark.Figure = bmp;
            bookmark.Text = Text;
            bookmark.Name = "";
            bookmark.Mapped = false;
            TempList.Add(bookmark);
            //Check if the FilterName is a preset in the document
            Check_Preset();
            //Map(MappedPos, TempList.Count - 1);
            lstBrowser_Populate(TempList);
        }

        /// <summary>
        ///The method that handles events triggered by the global event 'SendSweepData'
        /// </summary>
        /// <param name="FilterName">Give the data a name so that it can be filtered</param>
        /// <param name="Text">Additional info</param>
        /// <param name="FreqStart">Start frequency of the sweep</param>
        /// <param name="FreqStop">Stop frequency of the sweep</param>
        /// <param name="StepSize">Step of each frequency of the sweep</param>
        private void Update_ReceivedSweepData(string FilterName, string Text, string FreqStart, string FreqStop, string StepSize)
        {
            Reporting_Word.Bookmarks bookmark;
            bookmark.FilterTag = FilterName;
            bookmark.Figure = null;
            bookmark.Text = FreqStart + "MHz - " + FreqStop + "MHz with a " + StepSize + "MHz step increment." + Text;
            bookmark.Name = "";
            bookmark.Mapped = false;
            TempList.Add(bookmark);
            //Check if the FilterName is a preset in the document
            Check_Preset();
            lstBrowser_Populate(TempList);
        }

        private void Check_Preset()
        {
            for (int i = 0; i < FinalList.Count; i++)
            {
                for (int j = 0; j<TempList.Count;j++)
                {
                    if (FinalList[i].Name.Contains(TempList[j].FilterTag))//if there is a match
                    {
                        if (TempList[j].Mapped != true)
                        Map(i, j);
                    }
                }
            }
        }

        /// <summary>
        /// Populates the Browser list view with all available bookmarks that have not yet been mapped
        /// </summary>
        /// <param name="BookmarkList">List of all the bookmarks</param>
        private void lstBrowser_Populate(List<Reporting_Word.Bookmarks> BookmarkList)
        {
            lstBrowser.Items.Clear();
            int i = 0;
            foreach (var bookmark in BookmarkList)
            {
                lstBrowser.Items.Add(bookmark.FilterTag);
            }
        }

        /// <summary>
        /// Prompts the user for a file location where they would like to save the document
        /// </summary>
        /// <returns>Returns the file location if successful, else it returns ""</returns>
        private string SaveDialog()
        {
            SaveFileDialog SaveFile = new SaveFileDialog();
            SaveFile.Filter = "Word Document | *.docx ";
            SaveFile.Title = "Save Report";
            if (SaveFile.ShowDialog() == DialogResult.OK)
            {
                return SaveFile.FileName;
            }
            else
                return "";
        }
        #endregion
        
    }
}
