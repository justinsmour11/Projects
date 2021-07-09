using GemBox.Presentation;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTGenerator
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            imageList1.Images.Clear();

            //string oURL = "https://cse.google.com/cse?cx=042af5b82d449d07a&q=";
            //string APIkey = "AIzaSyC-Ccv1tnyYSq5RcQfpQLy77vfqBbNj52I";

            //string images = "http://images.google.com/images?q=";

            string bURL = "https://api.bing.microsoft.com/v7.0/search?q=";
            string bKey = "598b3e55b42e4c7aac82437cf189b080";
            

            //Bing request
            WebRequest request = WebRequest.Create(bURL + textBox1.Text);
            request.Headers["Ocp-Apim-Subscription-Key"] = bKey;

            WebResponse response = request.GetResponse();

            List<string> itemList = new List<string>();


            string responseFromServer;
            using (Stream dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                responseFromServer = reader.ReadToEnd();
            }

            response.Close();
            JObject jo = JObject.Parse(responseFromServer);
            try
            {
                JToken jItems = jo["images"]["value"];



            
                for (int i = 0; i < 6; i++)
                {
                    itemList.Add(jItems[i]["thumbnailUrl"].ToString());

                }
            }

            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine(ex.Message);
            }

            catch (ArgumentNullException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (NullReferenceException ex)
            {

            }

            //Google request
            string URL = "https://customsearch.googleapis.com/customsearch/v1?key=AIzaSyC-Ccv1tnyYSq5RcQfpQLy77vfqBbNj52I&cx=042af5b82d449d07a&q=";

            try
            {
                WebRequest gRequest = WebRequest.Create(URL + richTextBox1.Text);
                WebResponse gResponse = gRequest.GetResponse();

                string googleResponseFromServer;

                using (Stream dataStream = gResponse.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    googleResponseFromServer = reader.ReadToEnd();
                }

                gResponse.Close();
                JObject googleJo = JObject.Parse(googleResponseFromServer);
                JToken jItemsGoogle = googleJo["items"];

                for (int i = 0; i < 6; i++)
                {
                    itemList.Add(jItemsGoogle[i]["pagemap"]["cse_thumbnail"][0]["src"].ToString());

                }

            }
            catch (Exception ex)
            {

            }

           
            




            imageList1.ImageSize = new Size(256, 160);
            imageList1.ColorDepth = ColorDepth.Depth32Bit;
            listView1.Clear();
            for (int j = 0; j < itemList.Count; j++)
            {
                WebClient wc = new WebClient();
                byte[] imageByte = wc.DownloadData(itemList[j]);
                MemoryStream stream = new MemoryStream(imageByte);

                Image im = Image.FromStream(stream);
                im.Save($"image{j}.png");
                imageList1.Images.Add(im);

                listView1.Items.Add("", j);
            }
            itemList.Clear();

            listView1.LargeImageList = imageList1;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //CREATE SLIDE 1ST OPTION

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var presentation = new PresentationDocument();

            var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

            var shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 1, 1, 32, 16, LengthUnit.Centimeter);
            shape.Format.Fill.SetSolid(GemBox.Presentation.Color.FromName(ColorName.White));
            var title = shape.Text.AddParagraph().AddRun(textBox1.Text);
            title.Format.Fill.SetSolid(GemBox.Presentation.Color.FromName(ColorName.Black));

            var run = shape.Text.AddParagraph().AddRun(richTextBox1.Text);
            run.Format.Fill.SetSolid(GemBox.Presentation.Color.FromName(ColorName.Black));

            Picture picture = null;

            

            for (int i = 0; i < imageList2.Images.Count; i++)
            {
                imageList2.Images[i].Save($"image{i}.png");
                using (var stream = File.OpenRead($"image{i}.png"))
                    picture = slide.Content.AddPicture(PictureContentType.Png, stream, (i + 1) * 7, 2, 6, 5, LengthUnit.Centimeter);
            }
            



            try
            {
                presentation.Save("Slide.pptx");

            }

            catch (Exception ex)
            {

            }
            

        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            //int x = listView1.FocusedItem.Index;

            //imageList2.ImageSize = new Size(256, 160);
            //imageList2.ColorDepth = ColorDepth.Depth32Bit;

            

            //imageList2.Images.Add(imageList1.Images[x]);

            ////listView1.Items.Remove(listView1.SelectedItems[listView1.FocusedItem.Index]);
            ////listView2.Items.Add(listView1.SelectedItems[listView1.FocusedItem.Index]);


            //listView2.Items.Add("", listView1.FocusedItem.Index);
            


            

            //listView2.LargeImageList = imageList2;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //ADD
            int x = listView1.FocusedItem.Index;
            imageList2.ImageSize = new Size(256, 160);
            imageList2.ColorDepth = ColorDepth.Depth32Bit;
            imageList2.Images.Add(imageList1.Images[x]);

            
            listView2.Items.Add("", listView2.Items.Count);

            

            listView2.LargeImageList = imageList2;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            //BOLD
            if (!richTextBox1.SelectionFont.Bold)
            {
                richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Bold);

            }

            else
            {
                richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Regular);
            }

            //not working



        }

        private void button5_Click(object sender, EventArgs e)
        {
            //DIALOG BOX
            if (fontDialog1.ShowDialog()==DialogResult.OK & !String.IsNullOrEmpty(richTextBox1.Text))
            {
                richTextBox1.SelectionFont = fontDialog1.Font;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            int x = listView2.FocusedItem.Index;
            listView2.Items.RemoveAt(x);
        }
    }
}
