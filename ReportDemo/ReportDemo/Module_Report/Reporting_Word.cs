using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Module_Report
{
    public class Reporting_Word
    {

        public struct Bookmarks
        {
            public string Name;
            public string FilterTag;
            public string Text;
            public Bitmap Figure;
            public bool Mapped;
        }



        public Reporting_Word()
        {

        }

        /// <summary>
        /// Creates a Novacode Picture from a local image file such as .PNP, .JPG etc
        /// </summary>
        /// <param name="document">Referce to the global document</param>
        /// <param name="Directory">File directory of the local image file</param>
        /// <param name="Height">Height of the picture</param>
        /// <param name="Width">Width of the picture</param>
        /// <returns>Returns a picture of type Novacode.Picture</returns>
        public Picture CreatePicture(ref DocX document, string Directory, int Height, int Width)
        {
            Novacode.Image image = document.AddImage(Directory);
            // Create a picture (A custom view of an Image).
            Picture picture = image.CreatePicture();

            picture.Height = Height;
            picture.Width = Width;

            return picture;
        }

        /// <summary>
        /// Creates a Novacode Picture from a Bitmap
        /// </summary>
        /// <param name="document">Referce to the document</param>
        /// <param name="bmp">Bitmap to be converted</param>
        /// <param name="Height">Height of the picture</param>
        /// <param name="Width">Width of the picture</param>
        /// <returns>Returns a picture of type Novacode.Picture</returns>
        public Picture CreatePicture(ref DocX document, Bitmap bmp, int Height, int Width)
        {
            System.Drawing.Image image = (System.Drawing.Image)bmp;

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);// Save your picture in a memory stream.
                //image.Save(ms, bmp.RawFormat);  
                ms.Seek(0, SeekOrigin.Begin);

                Novacode.Image NovaImage = document.AddImage(ms); // Create image.

                Picture picture = NovaImage.CreatePicture();     // Create picture.

                picture.Height = Height;
                picture.Width = Width;

                return picture;
            }
        }

        /// <summary>
        /// Saves the document as a file on your computer
        /// </summary>
        /// <param name="document">The document you want to write to a file</param>
        /// <param name="FilePath">The location to save the file (C:\Document.docx) </param>
        public void SaveDoc(DocX document, string FilePath)
        {
            document.SaveAs(FilePath);
        }

        /// <summary>
        /// Loads an existing word (.docx) document
        /// </summary>
        /// <returns>Returns the opened document</returns>
        public DocX Load_Template()
        {
            OpenFileDialog LoadDoc = new OpenFileDialog();
            LoadDoc.Title = "Load document template";
            LoadDoc.Filter = "Word document |*.docx|All Supported Files |*.docx|All files|*.*";
            LoadDoc.FilterIndex = 1;    //By default the .docx file is selected
            LoadDoc.Multiselect = false;

            if (LoadDoc.ShowDialog() == DialogResult.OK)
            {
                DocX document = DocX.Load(LoadDoc.FileName);
                return document;
            }
            else return null;
        }

        /// <summary>
        /// Loads an existing word (.docx) document
        /// </summary>
        /// <param name="FilePath">The file path of the document e.g. C:\document.docx</param>
        /// <returns>Returns the opened document</returns>
        public DocX Load_Template(string FilePath)
        {
            DocX document = DocX.Load(FilePath);
            return document;
        }

        /// <summary>
        /// Gets a list of all the 
        /// </summary>
        /// <param name="document">Referce to the document</param>
        /// <returns>Returns a list of type string containing the names of all available bookmarks</returns>
        public List<Bookmarks> List_Bookmarks(DocX document)
        {
            List<Bookmarks> BookmarkList = new List<Bookmarks>();
            Bookmarks NewBookmark;
            foreach (Bookmark bookmark in document.Bookmarks)
            {
                //Perform a check to make sure it only adds applicable bookmarks (there are lots of random ones such as _Toc34827482 which we don't want)
                if (bookmark.Name.Contains("bmk"))
                {
                NewBookmark.Name = bookmark.Name;
                NewBookmark.Text = "";
                NewBookmark.Figure = null;
                NewBookmark.FilterTag = "";
                NewBookmark.Mapped = false;
                BookmarkList.Add(NewBookmark);
                Console.WriteLine("\t\tFound bookmark {0}", bookmark.Name);
                }
            }
            return BookmarkList;
        }

        /// <summary>
        /// Inserts a picture by the selected bookmark
        /// </summary>
        /// <param name="document">Referce to the document</param>
        /// <param name="bookmark">The exact name of the bookmark you want to change</param>
        /// <param name="picture">The picture you want to insert</param>
        /// <returns>Returns the modified document</returns>
        public DocX Bookmark_Change(DocX document, string bookmark, Picture picture)
        {
            document.Bookmarks[bookmark].Paragraph.AppendLine("").AppendPicture(picture);
            return document;
        }

        /// <summary>
        /// Inserts text by the selected bookmark
        /// </summary>
        /// <param name="document">Referce to the document</param>
        /// <param name="bookmark">The exact name of the bookmark you want to change</param>
        /// <param name="text">The new text of the selected bookmark</param>
        /// <returns>Returns the modified document</returns>
        public DocX Bookmark_Change(DocX document, string bookmark, string text)
        {
            document.Bookmarks[bookmark].SetText(text);
            return document;
        }

    }
}
