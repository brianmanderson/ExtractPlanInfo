using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.10")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace ExtractPlanInfo
{
  class Program
  {
    [STAThread]
    static void Main(string[] args)
    {
      try
      {
        using (Application app = Application.CreateApplication())
        {
          Execute(app);
        }
      }
      catch (Exception e)
      {
        Console.Error.WriteLine(e.ToString());
      }
    }
        static Tuple<HashSet<string>, HashSet<string>> ReadCSV(string csv_path)
        {
            HashSet<string> mrnList = new HashSet<string>();
            HashSet<string> diagnosisList = new HashSet<string>();
            using (var reader = new StreamReader(csv_path))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    mrnList.Add(values[0]);
                    diagnosisList.Add(values[1]);
                }
            }
            return Tuple.Create(mrnList, diagnosisList);
        }
    static StreamWriter return_streamwriter(string file_path, bool new_file)
        {
            if (!File.Exists(file_path) | (new_file))
            {
                StreamWriter fid_overallstatus = File.CreateText(file_path);
                return fid_overallstatus;
            }
            else
            {
                StreamWriter fid_overallstatus = File.AppendText(file_path);
                return fid_overallstatus;
            }
        }
    static void Execute(Application app)
    {
            // TODO: Add your code here.
            List<string> mrnList;
            List<string> diagnosisList;
            string patient_MRN;
            string out_path, patient_path;
            string base_path = @"C:\Users\b5anderson\Desktop\Plan_Data";
            string overall_path = Path.Combine(base_path, "All_Patients.txt");
            StreamWriter fid_overall = return_streamwriter(overall_path, false);
            string top_row = "MRN, CourseID, PlanID, BeamID, EnergyDisplayName, SSD, GantryAngle, Diagnosis Code";
            fid_overall.WriteLine(top_row);
            fid_overall.Close();
            var items = ReadCSV(@"K:\MRN_Diagnosis.csv");
            mrnList = items.Item1.ToList();
            diagnosisList = items.Item2.ToList();
            for (int i=1; i < mrnList.Count; i++)
            {
                patient_MRN = mrnList[i];
                patient_path = Path.Combine(base_path, patient_MRN);
                System.Console.WriteLine($"{patient_MRN}");
                if (Directory.Exists(patient_path)) // If this path already exists, just move along
                {
                    continue;
                }
                System.Threading.Thread.Sleep(1000);
                Patient pat = app.OpenPatientById(patient_MRN);
                if (pat is null)
                {
                    continue;
                }
                Directory.CreateDirectory(patient_path);
                foreach (Course course in pat.Courses)
                {
                    // Check to see if the diagnosis from our excel sheet is present here
                    bool has_diagnosis = false;
                    foreach (Diagnosis diagnosis in course.Diagnoses)
                    {
                        if (diagnosisList.Contains(diagnosis.Code.Trim()))
                        {
                            has_diagnosis = true;
                            out_path = Path.Combine(patient_path, $"{course.Id.Replace('/','.').Replace(':','.')}.txt");
                            System.Console.WriteLine($"{out_path}");
                            StreamWriter fid = return_streamwriter(out_path, true);
                            fid.WriteLine(top_row);
                            // Lets pull gantry angle, SSD, and bolus info
                            foreach (ExternalPlanSetup external_beam_plan in course.ExternalPlanSetups)
                            {
                                foreach (Beam beam in external_beam_plan.Beams)
                                {
                                    double SSD = beam.PlannedSSD;
                                    string info = $"{mrnList[i]},{course.Id.Replace('/', '.')},{external_beam_plan.Id},{beam.Id},{beam.EnergyModeDisplayName},{SSD},{beam.ControlPoints[0].GantryAngle},{diagnosis.Code.Trim()}";
                                    fid.WriteLine(info);
                                    StreamWriter fid_over = return_streamwriter(overall_path, false);
                                    fid_over.WriteLine(info);
                                    fid_over.Close();
                                    ControlPointCollection controlpoints = beam.ControlPoints;
                                }
                            }
                            fid.Close();
                            break;
                        }
                    }
                    if (!has_diagnosis)
                    {
                        continue;
                    }
                }
                app.ClosePatient();
            }

    }
  }
}
