/*                                                                    * Auther :Ahmed Ali Abd al mola*
                                                                          *  creat in :22/1/2018*
                                                                            *                  *
 *                                                                           ******************
 *                                                                            ****************
 *                                                                             **************
 *                                                                              ************
 * */

using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Messaging;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Threading;
using System.Dynamic;
using System.Drawing.Drawing2D;
using Excel = Microsoft.Office.Interop.Excel;  //the new liprary after add refrance microsoft Excel 14 object in references-> COM tab

namespace WindowsFormsApplicationfirst
{ public partial class Form1 : Form
{
   

    public Form1() { InitializeComponent(); }


    private void Form1_Load(object sender, EventArgs e)
    {
       
        }


         public int numpro, numchose=0;
  
    //--------------------------------------------------------------------------------------------   
    string afteredit, beforedit, textpath, cp;
    int rowCount;                                       //           viriable
    int colCount;
    bool star = false, finsh = false, gptr = true,FRR=false;
    int ee = 2;

     //-------------------------------------------------------------------------------------------
        //--------------------------------
         //Read_From_Excel rfe = new Read_From_Excel();
     //----------------------------------------------------------------------------------------------------
        private void pictureBox1_Click(object sender, EventArgs e){
            Process.Start("http://www.aast.edu/en/");
        }   //Internet go
        private void pictureBox4_Click(object sender, EventArgs e) {
            Process.Start("https://www.facebook.com/RIC.AAST.Aswan/");
        }
        //----------------------------------------------------------------------------------------------------


        
         
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            int locatinmark = 0, lenthofpath = 0;

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Exel file (*.xlsx*)|*.xlsx*";
            choofdlog.FilterIndex = 1;
            
            choofdlog.Multiselect = true;
            choofdlog.ShowDialog(); afteredit= choofdlog.ToString();
          //  textBox1.Text= choofdlog.ToString();
            afteredit = afteredit.Substring(56, afteredit.Length - 56);
            beforedit = afteredit;
            textBox1.Text = afteredit;
         //--------------------------------------------------------------------------
            if (afteredit.Length > 1)
            {
                for (lenthofpath = (afteredit.Length - 1); lenthofpath >= 0; lenthofpath--)
                {
                    if (afteredit[lenthofpath] == (char)'\\') { locatinmark = lenthofpath; break; }
                }
                //  text addres
                textpath = afteredit.Substring(0, locatinmark + 1);
               cp = textpath;
                textpath = textpath + "The Report.txt";
            
           //---------------------------------------------------------------------------- 
          
                MessageBox.Show("!!اذا كان الملف منظم بشكل غير الموجود ف النموزج سوف يحدث خطأ ");
            
                label14.Text = "Add Done";
                FRR = true;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@afteredit);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                rowCount = xlRange.Rows.Count;
                colCount = xlRange.Columns.Count;


                label2.Text = (rowCount - 1).ToString();
                label3.Text = (colCount - 4).ToString();



                Console.Write(xlRange.Rows.Count);

                Console.Write("\n" + e);
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
 }
     
        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (FRR)
            {
                string nuu = textBox2.Text;
                Boolean ra = false;
                for (int i = 0; i < nuu.Length; i++)
                    if (nuu == "0" || (nuu[i] > (char)57) || (nuu[i] < (char)48))
                    { ra = true; break; }


                if (textBox2.Text != "" && !ra)
                {



                    numpro = Convert.ToInt16(textBox2.Text);
                    pictureBox7.Size = new Size(0, 0);
                    label14.Text = "Enter Done";
                    star = true;
                }
                else if (ra)
                    MessageBox.Show("!!يوجد خطأ ف الادخال  ");
                else { MessageBox.Show("!!االرجاء ملءالخانات  "); }
            }
            else MessageBox.Show("!!االرجاء اختيار الملف  ");
           
        }
        
        private void button1_Click(object sender, EventArgs e)  


        {



           


            try
            {




                if (finsh == true)
                {
                    Environment.Exit(0);
                }

                else if (star&&FRR)
                {
                    File.Copy(@beforedit, @cp + "yourcopy.xlsx");
                    //--------------------------------------------------------------------------------------------
                    int i, j, k, l, op, donecount = 0, fakecount = 0, notdonecount = 0, projectnum = 0;//           viriable
                    double lastgpa = 0, firstgpa = 0;
                    //-------------------------------------------------------------------------------------------
                    string path = @textpath;
                    if (!File.Exists(path)) using (StreamWriter sw = File.CreateText(path))
                        {  //                                                                        file txte ceriate
                            sw.WriteLine("______________________________________________________________________________________________________________________________________________________________");
                            sw.WriteLine("                                                                        welcome in Project Organizer");
                            sw.WriteLine("______________________________________________________________________________________________________________________________________________________________");
                        }
                    //--------------------------------------------------------------------------------------------------------

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@beforedit);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    projectnum = colCount - 4;  //                                                             project number*****
                    //-----------------------------------------------------------------------------------------------
                    string[,] mem = new string[rowCount + 1, colCount + 1];
                    for (i = 2; i <= rowCount; i++)
                    {
                        for (j = 1; j <= colCount; j++)
                        {                                     //                                              memorization*******
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                mem[i, j] = xlRange.Cells[i, j].Value2.ToString();
                            }
                            else mem[i, j] = "";
                        }
                    }
                    for (j = 5; j <= colCount; j++)
                        if (mem[3, j] != null && mem[3, j] != "")
                            numchose++;

                    //--------------------------------------------------------------- s o r t
                    //--------------------------------------------------------------  s o r t
                    for (i = 2; i <= rowCount; i++)
                        
                    {
                        Double g;
                        String v;
                        v = mem[i, 4];
                        g = Convert.ToDouble(v);

                        if (g > 4)
                        {
                            gptr = false;
                            MessageBox.Show("يوجد شخص قام ب ادخال معدل تراكمى اكبرمن المتفق عليه");

                            //cleanup
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                            //rule of thumb for releasing com objects:
                            //  never use two dots, all COM objects must be referenced and released individually
                            //  ex: [somthing].[something].[something] is bad

                            //release com objects to fully kill excel process from running in the background
                            Marshal.ReleaseComObject(xlRange);
                            Marshal.ReleaseComObject(xlWorksheet);
                            xlWorkbook.Save();

                            //close and release
                            xlWorkbook.Close();
                            Marshal.ReleaseComObject(xlWorkbook);

                            //quit and release





                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlApp);
                        }
                    }
                    if (gptr)
                    {
                        Double g1, g2;
                        int[] stu = new int[102];
                        string ch1;

                        for (j = 2; j <= rowCount; j++)
                        {
                            for (i = j + 1; i <= rowCount; i++)
                            {
                                String v1, v2;
                                v1 = mem[j, 4];
                                g1 = Convert.ToDouble(v1);
                                v2 = mem[i, 4];
                                g2 = Convert.ToDouble(v2);
                                if (g1 < g2)
                                {

                                    for (l = 1; l <= colCount; l++)
                                    {

                                        if ((((mem[j, l] == null || mem[j, l] == ""))) && (((mem[i, l] == null || mem[i, l] == ""))))
                                        {

                                            mem[j, l] = "";
                                            mem[i, l] = "";
                                        }


                                        else if ((((mem[j, l] == null || mem[j, l] == ""))) && (((mem[i, l] != null && mem[i, l] != ""))))
                                        {

                                            ch1 = mem[i, l];
                                            mem[j, l] = ch1;
                                            mem[i, l] = "";
                                        }


                                        else if ((((mem[j, l] != null && mem[j, l] != ""))) && (((mem[i, l] == null || mem[i, l] == ""))))
                                        {

                                            ch1 = mem[j, l];
                                            mem[i, l] = ch1;
                                            mem[j, l] = "";
                                        }


                                        else if ((((mem[j, l] != null && mem[j, l] != ""))) && (((mem[i, l] != null && mem[i, l] != ""))))
                                        {

                                            ch1 = mem[j, l];
                                            mem[j, l] = mem[i, l];
                                            mem[i, l] = ch1;
                                        }
                                    }
                                }
                            }
                        }
                        pictureBox9.Size = new Size(0, 0);
                        label14.Text = "Sort Done";
                        this.ResumeLayout();
                        //----------------------------------------------------------------------------------------------------
                        //---------------------------------------------------------------------------------------------------
                        using (System.IO.StreamWriter file =
                      new System.IO.StreamWriter(@textpath, true))
                        {
                            file.WriteLine(" \n");
                            file.WriteLine("                                                                                   (1):Sort is Done");   //                                                  2222222
                            file.WriteLine("                                                           -------------------------------------------------------------------------\n");
                        }
                        //---------------------------------------------------------------------------------------------------
                        //----------------------------------------------------------------------------------------------------
                        using (System.IO.StreamWriter file =
                      new System.IO.StreamWriter(@textpath, true))
                        {
                            file.WriteLine("                                                                                          (2):Fake Main:\n");
                            file.WriteLine("                                                           -------------------------------------------------------------------------\n");
                            file.WriteLine(" \n");
                        }
                        //-----------------------------------------------------------------------------------------------------
                        //------------------------------------------------------------------------------------------------------
                        for (i = 2; i < rowCount; i++)
                        {
                            for (j = i + 1; j <= rowCount; j++)
                            {
                                if ((mem[i, 3] != null && mem[i, 3] != "") && (mem[j, 3] != null && mem[j, 3] != ""))
                                {
                                    if ((mem[j, 3] == mem[i, 3]))
                                    {
                                        mem[j, 1] = " Fake ";
                                        fakecount++;
                                        //----------------------------------------------------------------------------------------------------
                                        using (System.IO.StreamWriter file =
                                      new System.IO.StreamWriter(@textpath, true))
                                        {
                                            file.WriteLine(mem[i, 2]);

                                            file.WriteLine("\n");
                                        }

                                        //------------------------------------------------------------------------------------------------------

                                        for (k = 2; k <= colCount; k++)
                                            mem[j, k] = "";
                                    }
                                }

                            }
                        }
                        pictureBox11.Size = new Size(0, 0);
                        label14.Text = "Fake Done";
                       

                        //-----------------------------------------------------------------------------------------------------
                        //------------------------------------------------------------------------------------------------------
                        //---------------------------------
                        
                        for (i = 0; i < 102; i++)
                            stu[i] = 0;             //          all array=0
                        // -----------------------------------
                        if (numchose > 0)
                        {
                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {
                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {


                                        if (mem[i, j] == "رغبة اولي" || mem[i, j] == "A")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first


                                            ee++;
                                            ++stu[j];
                                            xlRange.Cells[(int)(rowCount + ee), (int)j] = mem[i, 2];
                                            donecount++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }


                                        }


                                    }


                                }
                            }
                        }

                        //---------------------------------------------------------------------------------------
                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة ثانية" || mem[i, j] == "B")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //------------------------------------------------------------
                        if (numchose > 0)
                        {
                            numchose--;
                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {




                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {


                                        if (mem[i, j] == "رغبة ثالثة" || mem[i, j] == "C")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first


                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            mem[(int)(i), (int)(1)] = " Done ";
                                            ee++;
                                            stu[j]++;
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }

                                        }

                                    }


                                }
                            }
                        }


                        //-------------------------------------------
                        //--------------------------------------------
                        //------------------------------------------
                        //------------------------------------------
                        //----------------------------------------

                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة رابعة" || mem[i, j] == "D")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }



                        //-------------------------------------------------------------

                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة خامسة" || mem[i, j] == "E")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //--------------------------------------------------------------------
                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة سادسة" || mem[i, j] == "F")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }


                        //------------------------------------------------------------------

                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة سابعة" || mem[i, j] == "G")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }



                        //----------------------------------------------------
                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة ثامنة" || mem[i, j] == "H")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }



                        //--------------------------------------------------------
                        if (numchose > 0)
                        {
                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة تاسعة" || mem[i, j] == "I")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }


                        //-------------------------------------------------------------------

                        if (numchose > 0)
                        {

                            numchose--;

                            for (j = 5; j <= colCount; j++)
                            {
                                ee = 2;

                                for (i = 2; i <= rowCount; i++)
                                {

                                    if (mem[i, j] != null && mem[i, j] != "")
                                    {
                                        if (mem[i, j] == "رغبة عاشرة" || mem[i, j] == "J")
                                        {
                                            //-----------------------------------------------------------------------
                                            firstgpa = Convert.ToDouble((mem[i, 4]));
                                            if ((stu[j] >= numpro) && (lastgpa != firstgpa))
                                            { lastgpa = 0; break; }
                                            lastgpa = firstgpa;
                                            //------------------------------------------------------------------last gpa and first
                                            donecount++;
                                            xlRange.Cells[rowCount + ee, j] = mem[i, 2];

                                            ee++;
                                            stu[j]++;
                                            mem[(int)(i), (int)(1)] = " Done ";
                                            for (op = 5; op <= colCount; op++)
                                            {
                                                mem[i, op] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }




                        notdonecount = rowCount - donecount - fakecount - 1;

                        //----------------------------------------------------------------------------------------------------
                        using (System.IO.StreamWriter file =
                      new System.IO.StreamWriter(@textpath, true))
                        {
                            file.WriteLine("                                                              -------------------------------------------------------------------------\n");
                            file.WriteLine("                                                                                          (3)chose  is Done:\n");
                            file.WriteLine("                                                              -------------------------------------------------------------------------\n");
                            file.WriteLine("                                                                                          [1]:the number of Project :" + Convert.ToString(projectnum));
                            file.WriteLine("                                                                                          [2]:the number of all Student :" + Convert.ToString(rowCount - 1));
                            file.WriteLine("                                                                                          [3]:the number of Not Done :" + Convert.ToString(notdonecount));
                            file.WriteLine("                                                                                          [4]:the number of  Done :" + Convert.ToString(donecount));
                            file.WriteLine("                                                                                          [5]:the number of Fake :" + Convert.ToString(fakecount));
                            file.WriteLine("\n");
                            file.WriteLine("\n");
                            file.WriteLine("\n");
                            file.WriteLine("\n");
                            file.WriteLine("--------------------------------------------------------------------------------------------");
                            file.WriteLine("\n");
                            file.WriteLine("                                                         المركز الاقليمى للمعلوماتية (اسوان)000 ");
                            file.WriteLine("\n");
                            file.WriteLine("--------------------------------------------------------------------------------------------");

                            label4.Text = Convert.ToString(fakecount);
                            label6.Text = Convert.ToString(donecount);
                            label7.Text = Convert.ToString(notdonecount);
                        }
                        //-----------------------------------------------------------------------------------------------------

                        // Console.Write(xlRange.Rows.Count);
                        label14.Text = "Shose Done";
                        pictureBox12.Size = new Size(0, 0);

                        for (i = 2; i <= rowCount; i++)
                        {
                            for (j = 1; j <= colCount; j++)
                            {
                                xlRange.Cells[i, j] = mem[i, j];
                            }
                        }



                        //cleanup
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        //rule of thumb for releasing com objects:
                        //  never use two dots, all COM objects must be referenced and released individually
                        //  ex: [somthing].[something].[something] is bad

                        //release com objects to fully kill excel process from running in the background
                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);
                        xlWorkbook.Save();

                        //close and release
                        xlWorkbook.Close();
                        Marshal.ReleaseComObject(xlWorkbook);

                        //quit and release





                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                        button1.Text = "Finish";
                        MessageBox.Show("تمت المهمة بنجاح.. تستطيع الان الخروج والاطلاع عل الملف والتقرير الخاص بة ");
                        
                        label14.Text = "complite";
                        finsh = true;
                    }
                    




                }
                else MessageBox.Show("الرجاء استكمال الخطوات ");
            }
            catch (Exception AhmedTopCoder)
            {
                MessageBox.Show("لقد حدث خطأ اثناء المعالجه \n:يرجى التأكد من نمذج الملف او البيانات المدخلة ");
                //MessageBox.Show(AhmedTopCoder.ToString());

            }

          

        }

        private void button4_Click(object sender, EventArgs e)
        {
       
            Environment.Exit(0);
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {

        }
      
        private void button3_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            frm.Show();

       }

        private void Form1_Resize(object sender, EventArgs e)
        {
           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int locatinmark = 0, lenthofpath = 0;

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Exel file (*.xlsx*)|*.xlsx*";
            choofdlog.FilterIndex = 1;

            choofdlog.Multiselect = true;
            choofdlog.ShowDialog(); afteredit = choofdlog.ToString();
            //  textBox1.Text= choofdlog.ToString();
            afteredit = afteredit.Substring(56, afteredit.Length - 56);
            beforedit = afteredit;
            textBox1.Text = afteredit;
            //--------------------------------------------------------------------------
            if (afteredit.Length > 1)
            {
                for (lenthofpath = (afteredit.Length - 1); lenthofpath >= 0; lenthofpath--)
                {
                    if (afteredit[lenthofpath] == (char)'\\') { locatinmark = lenthofpath; break; }
                }
                //  text addres
                textpath = afteredit.Substring(0, locatinmark + 1);
                cp = textpath;
                textpath = textpath + "The Report.txt";

                //---------------------------------------------------------------------------- 

                MessageBox.Show("!!اذا كان الملف منظم بشكل غير الموجود ف النموزج سوف يحدث خطأ ");

                label14.Text = "Add Done"; FRR = true;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@afteredit);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                rowCount = xlRange.Rows.Count;
                colCount = xlRange.Columns.Count;


                label2.Text = (rowCount - 1).ToString();
                label3.Text = (colCount - 4).ToString();



                Console.Write(xlRange.Rows.Count);

                Console.Write("\n" + e);
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void pictureBox11_Click_1(object sender, EventArgs e)
        {

        }
}

    

    
 } 