using IronXL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp9
{
    public partial class Form1 : Form
    {
        List<string> numere = new List<string>();
        List<string> fisiereDeScris = new List<string>();
        List<string> rangeDeCitit = new List<string>();
        

    string st10FlowVolumeRange = Properties.Settings.Default.st10FlowVolumeRange;
        string st10WeldingDistanceRange = Properties.Settings.Default.st10WeldingDistanceRange;
        string st40_point1_Range = Properties.Settings.Default.st40_point1_Range;
        string st40_point2_Range = Properties.Settings.Default.st40_point2_Range;
        string st40_point3_Range = Properties.Settings.Default.st40_point3_Range;
        string st40_point4_Range = Properties.Settings.Default.st40_point4_Range;
        string st40_point5_Range = Properties.Settings.Default.st40_point5_Range;
        string st40_point6_Range = Properties.Settings.Default.st40_point6_Range;
        string st60_welding_distance_Range = Properties.Settings.Default.st60_welding_distance_Range;
        string st70_plug_force_Range = Properties.Settings.Default.st70_plug_force_Range;
        string st80_insulation_resistance_hv_Range = Properties.Settings.Default.st80_insulation_resistance_hv_Range;
        string st80_leak_test_Range = Properties.Settings.Default.st80_leak_test_Range;
        string st80_resistivity_cp_Range = Properties.Settings.Default.st80_resistivity_cp_Range;
        string st80_resistivity_pe_to_cp_Range = Properties.Settings.Default.st80_resistivity_pe_to_cp_Range;
        string st80_resistivity_pp_Range = Properties.Settings.Default.st80_resistivity_pp_Range;
        string st80_temperature_dc_minus_Range = Properties.Settings.Default.st80_temperature_dc_minus_Range;
        string st80_temperature_dc_plus_Range = Properties.Settings.Default.st80_temperature_dc_plus_Range;
        string st90_pin_length_1_Range = Properties.Settings.Default.st90_pin_length_1_Range;
        string st90_pin_length_2_Range = Properties.Settings.Default.st90_pin_length_2_Range;
        string st90_pin_length_3_Range = Properties.Settings.Default.st90_pin_length_3_Range;
        string st90_pin_length_5_Range = Properties.Settings.Default.st90_pin_length_5_Range;
        string st90_pin_length_6_Range = Properties.Settings.Default.st90_pin_length_6_Range;
        string st90_pin_length_7_Range = Properties.Settings.Default.st90_pin_length_7_Range;
        string st90_pin_length_8_Range = Properties.Settings.Default.st90_pin_length_8_Range;


        public Form1()
        {
            InitializeComponent();
            
            SetareValoriCampuriFisiere();
            
            SetareValoriCampuriRange();

            InitializareFisiereDeScris();

            InitializareRange();  

        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool[] ss = { false, false, false, false, false, false, false, false, false, false,
                          false, false,false, false,false, false,false, false,false, false,false, false,
                          false, false,false
                        };
            ss[0] = chk_st10_flowVolume.Checked;
            ss[1] = chk_st10_weldingDistance.Checked;
            ss[2] = chk_st40_point1.Checked;
            ss[3] = chk_st40_point2.Checked; 
            ss[4] = chk_st40_point3.Checked; 
            ss[5] = chk_st40_point4.Checked; 
            ss[6] = chk_st40_point5.Checked; 
            ss[7] = chk_st40_point6.Checked;
            ss[8] = chk_st60_welding_distance.Checked;
            ss[9] = chk_st70_plug_force.Checked;
            ss[10] = chk_st80_insulation_resistance_hv.Checked;
            ss[11] = chk_st80_leak_test.Checked;
            ss[12] = chk_st80_resistivity_cp.Checked;
            ss[13] = chk_st80_resistivity_pe_to_cp.Checked;
            ss[14] = chk_st80_resistivity_pp.Checked;
            ss[15] = chk_st80_temperature_dc_minus.Checked;
            ss[16] = chk_st80_temperature_dc_plus.Checked;
            ss[17] = chk_st90_pin_length_1.Checked;
            ss[18] = chk_st90_pin_length_2.Checked;
            ss[19] = chk_st90_pin_length_3.Checked;
            ss[20] = chk_st90_pin_length_4.Checked;
            ss[21] = chk_st90_pin_length_5.Checked;
            ss[22] = chk_st90_pin_length_6.Checked;
            ss[23] = chk_st90_pin_length_7.Checked;
            ss[24] = chk_st90_pin_length_8.Checked;


            for (int i = 0; i < fisiereDeScris.Count; i++)
            {
                if (ss[i]== true)
                {
                    continue;
                }
                else
                {
                    ReadExcelFile(rangeDeCitit[i]);
                    WriteExcelFile(fisiereDeScris[i]);
                    numere.Clear();
                } 
            }
           
        }

        public void ReadExcelFile(string rangeDeCitit)
        {
            WorkBook workbook = WorkBook.Load(txt_logfile.Text);
            WorkSheet sheet = workbook.WorkSheets.First();

            //Citirea randurilor dintr-un anumit interval de pe o coloana
            var range = sheet[rangeDeCitit];
            //Citirea valorilor din fiecare celula si introducerea lor in lista
            foreach (var cell in range)
            {
                numere.Add(cell.Value.ToString());
            }
        }

        public void WriteExcelFile(string deScris)
        {
            WorkBook workbook2 = WorkBook.Load(deScris);
            WorkSheet sheet2 = workbook2.DefaultWorkSheet;

            for (int i = 0; i < numere.Count; i++)
            {
                sheet2["C" + (i + 5)].Value = numere[i];
                sheet2["C" + (i + 5)].First().FormatString = "0.000";
                //  sheet2["C" + (i + 5)].NumberFormat = "##.000";
            }
            //Save Changes
            workbook2.SaveAs(deScris);
        }

        public void InitializareFisiereDeScris()
        {
            fisiereDeScris.Add(txt_st10_flowVolume.Text);
            fisiereDeScris.Add(txt_st10_weldingDistance.Text);
            fisiereDeScris.Add(txt_st40_point1.Text);
            fisiereDeScris.Add(txt_st40_point2.Text);
            fisiereDeScris.Add(txt_st40_point3.Text);
            fisiereDeScris.Add(txt_st40_point4.Text);
            fisiereDeScris.Add(txt_st40_point5.Text);
            fisiereDeScris.Add(txt_st40_point6.Text);
            fisiereDeScris.Add(txt_st60_welding_distance.Text);
            fisiereDeScris.Add(txt_st70_plug_force.Text);
            fisiereDeScris.Add(txt_st80_insulation_resistance_hv.Text);
            fisiereDeScris.Add(txt_st80_leak_test.Text);
            fisiereDeScris.Add(txt_st80_resistivity_cp.Text);
            fisiereDeScris.Add(txt_st80_resistivity_pe_to_cp.Text);
            fisiereDeScris.Add(txt_st80_resistivity_pp.Text);
            fisiereDeScris.Add(txt_st80_temperature_dc_minus.Text);
            fisiereDeScris.Add(txt_st80_temperature_dc_plus.Text);
            fisiereDeScris.Add(txt_st90_pin_length_1.Text);
            fisiereDeScris.Add(txt_st90_pin_length_2.Text);
            fisiereDeScris.Add(txt_st90_pin_length_3.Text);
            fisiereDeScris.Add(txt_st90_pin_length_5.Text);
            fisiereDeScris.Add(txt_st90_pin_length_6.Text);
            fisiereDeScris.Add(txt_st90_pin_length_7.Text);
            fisiereDeScris.Add(txt_st90_pin_length_8.Text);
        }
        public void InitializareRange()
        {
            rangeDeCitit.Add(st10FlowVolumeRange);
            rangeDeCitit.Add(st10WeldingDistanceRange);
            rangeDeCitit.Add(st40_point1_Range);
            rangeDeCitit.Add(st40_point2_Range);
            rangeDeCitit.Add(st40_point3_Range);
            rangeDeCitit.Add(st40_point4_Range);
            rangeDeCitit.Add(st40_point5_Range);
            rangeDeCitit.Add(st40_point6_Range);
            rangeDeCitit.Add(st60_welding_distance_Range);
            rangeDeCitit.Add(st70_plug_force_Range);
            rangeDeCitit.Add(st80_insulation_resistance_hv_Range);
            rangeDeCitit.Add(st80_leak_test_Range);
            rangeDeCitit.Add(st80_resistivity_cp_Range);
            rangeDeCitit.Add(st80_resistivity_pe_to_cp_Range);
            rangeDeCitit.Add(st80_resistivity_pp_Range);
            rangeDeCitit.Add(st80_temperature_dc_minus_Range);
            rangeDeCitit.Add(st80_temperature_dc_plus_Range);
            rangeDeCitit.Add(st90_pin_length_1_Range);
            rangeDeCitit.Add(st90_pin_length_2_Range);
            rangeDeCitit.Add(st90_pin_length_3_Range);
            rangeDeCitit.Add(st90_pin_length_5_Range);
            rangeDeCitit.Add(st90_pin_length_6_Range);
            rangeDeCitit.Add(st90_pin_length_7_Range);
            rangeDeCitit.Add(st90_pin_length_8_Range);
        }

       public void SetareValoriCampuriRange()
        {
            txt_st10_flowVolume_Range.Text = st10FlowVolumeRange;
            txt_st10_weldingDistance_Range.Text = st10WeldingDistanceRange;
            txt_st40_point1_Range.Text = st40_point1_Range;
            txt_st40_point2_Range.Text = st40_point2_Range;
            txt_st40_point3_Range.Text = st40_point3_Range;
            txt_st40_point4_Range.Text = st40_point4_Range;
            txt_st40_point5_Range.Text = st40_point5_Range;
            txt_st40_point6_Range.Text = st40_point6_Range;
            txt_st60_welding_distance_Range.Text = st60_welding_distance_Range;
            txt_st70_plug_force_Range.Text = st70_plug_force_Range;
            txt_st80_insulation_resistance_hv_Range.Text = st80_insulation_resistance_hv_Range;
            txt_st80_leak_test_Range.Text = st80_leak_test_Range;
            txt_st80_resistivity_cp_Range.Text = st80_resistivity_cp_Range;
            txt_st80_resistivity_pe_to_cp_Range.Text = st80_resistivity_pe_to_cp_Range;
            txt_st80_resistivity_pp_Range.Text = st80_resistivity_pp_Range;
            txt_st80_temperature_dc_minus_Range.Text = st80_temperature_dc_minus_Range;
            txt_st80_temperature_dc_plus_Range.Text = st80_temperature_dc_plus_Range;
            txt_st90_pin_length_1_Range.Text = st90_pin_length_1_Range;
            txt_st90_pin_length_2_Range.Text = st90_pin_length_2_Range;
            txt_st90_pin_length_3_Range.Text = st90_pin_length_3_Range;
            txt_st90_pin_length_5_Range.Text = st90_pin_length_5_Range;
            txt_st90_pin_length_6_Range.Text = st90_pin_length_6_Range;
            txt_st90_pin_length_7_Range.Text = st90_pin_length_7_Range;
            txt_st90_pin_length_8_Range.Text = st90_pin_length_8_Range;

        }
        public void SetareValoriCampuriFisiere()
        {
            txt_logfile.Text = Properties.Settings.Default.txt_logfile;
            txt_st10_flowVolume.Text = Properties.Settings.Default.txt_st10_flowVolume;
            txt_st10_weldingDistance.Text = Properties.Settings.Default.txt_st10_weldingDistance;
            txt_st40_point1.Text = Properties.Settings.Default.txt_st40_point1;
            txt_st40_point2.Text = Properties.Settings.Default.txt_st40_point2;
            txt_st40_point3.Text = Properties.Settings.Default.txt_st40_point3;
            txt_st40_point4.Text = Properties.Settings.Default.txt_st40_point4;
            txt_st40_point5.Text = Properties.Settings.Default.txt_st40_point5;
            txt_st40_point6.Text = Properties.Settings.Default.txt_st40_point6;
            txt_st60_welding_distance.Text = Properties.Settings.Default.txt_st60_welding_distance;
            txt_st70_plug_force.Text = Properties.Settings.Default.txt_st70_plug_force;
            txt_st80_insulation_resistance_hv.Text = Properties.Settings.Default.txt_st80_insulation_resistance_hv;
            txt_st80_leak_test.Text = Properties.Settings.Default.txt_st80_leak_test;
            txt_st80_resistivity_cp.Text = Properties.Settings.Default.txt_st80_resistivity_cp;
            txt_st80_resistivity_pe_to_cp.Text = Properties.Settings.Default.txt_st80_resistivity_pe_to_cp;
            txt_st80_resistivity_pp.Text = Properties.Settings.Default.txt_st80_resistivity_pp;
            txt_st80_temperature_dc_minus.Text = Properties.Settings.Default.txt_st80_temperature_dc_minus;
            txt_st80_temperature_dc_plus.Text = Properties.Settings.Default.txt_st80_temperature_dc_plus;
            txt_st90_pin_length_1.Text = Properties.Settings.Default.txt_st90_pin_length_1;
            txt_st90_pin_length_2.Text = Properties.Settings.Default.txt_st90_pin_length_2;
            txt_st90_pin_length_3.Text = Properties.Settings.Default.txt_st90_pin_length_3;
            txt_st90_pin_length_5.Text = Properties.Settings.Default.txt_st90_pin_length_5;
            txt_st90_pin_length_6.Text = Properties.Settings.Default.txt_st90_pin_length_6;
            txt_st90_pin_length_7.Text = Properties.Settings.Default.txt_st90_pin_length_7;
            txt_st90_pin_length_8.Text = Properties.Settings.Default.txt_st90_pin_length_8;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.txt_logfile = txt_logfile.Text;
            Properties.Settings.Default.txt_st10_flowVolume = txt_st10_flowVolume.Text;
            Properties.Settings.Default.txt_st10_weldingDistance = txt_st10_weldingDistance.Text;
            Properties.Settings.Default.txt_st40_point1 = txt_st40_point1.Text;
            Properties.Settings.Default.txt_st40_point2 = txt_st40_point2.Text;
            Properties.Settings.Default.txt_st40_point3 = txt_st40_point3.Text;
            Properties.Settings.Default.txt_st40_point4 = txt_st40_point4.Text;
            Properties.Settings.Default.txt_st40_point5 = txt_st40_point5.Text;
            Properties.Settings.Default.txt_st40_point6 = txt_st40_point6.Text;
            Properties.Settings.Default.txt_st60_welding_distance = txt_st60_welding_distance.Text;
            Properties.Settings.Default.txt_st70_plug_force = txt_st70_plug_force.Text;
            Properties.Settings.Default.txt_st80_insulation_resistance_hv = txt_st80_insulation_resistance_hv.Text;
            Properties.Settings.Default.txt_st80_leak_test = txt_st80_leak_test.Text;
            Properties.Settings.Default.txt_st80_resistivity_cp = txt_st80_resistivity_cp.Text;
            Properties.Settings.Default.txt_st80_resistivity_pe_to_cp = txt_st80_resistivity_pe_to_cp.Text;
            Properties.Settings.Default.txt_st80_resistivity_pp = txt_st80_resistivity_pp.Text;
            Properties.Settings.Default.txt_st80_temperature_dc_minus = txt_st80_temperature_dc_minus.Text;
            Properties.Settings.Default.txt_st80_temperature_dc_plus = txt_st80_temperature_dc_plus.Text;
            Properties.Settings.Default.txt_st90_pin_length_1 = txt_st90_pin_length_1.Text;
            Properties.Settings.Default.txt_st90_pin_length_2 = txt_st90_pin_length_2.Text;
            Properties.Settings.Default.txt_st90_pin_length_3 = txt_st90_pin_length_3.Text;
            Properties.Settings.Default.txt_st90_pin_length_5 = txt_st90_pin_length_5.Text;
            Properties.Settings.Default.txt_st90_pin_length_6 = txt_st90_pin_length_6.Text;
            Properties.Settings.Default.txt_st90_pin_length_7 = txt_st90_pin_length_7.Text;
            Properties.Settings.Default.txt_st90_pin_length_8= txt_st90_pin_length_8.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Modificarile au fost salvate!");
        }
       
    }
}
