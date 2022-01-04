using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace TestTangible
{
    public class DataGenerator
    {
        public static DataTable GenerateExampleTableOne( )
        {
            var colors = new List<string>() {
            "661 red",
            "595 orange",
            "595 orange",
            "575 yellow",
            "523 green",
            "523 green",
            "482 blue",
            "450 indigo",
            "450 indigo",
            "423 violet" };
            var gas_molforms = new List<string>() {
                "CH4",
                "C3F6",
                "C2H6",
                "Ar",
                "H2O",
                "CF4",
                "SF6"
            };
            var table = DataGenerator.LightIntensityTable(Enumerable.Range(4567, 7 * 8),
                colors, gas_molforms);
            return table;
        }

        public static string MakeRandomString(int length)
        {
            var random_generator = new System.Random();
            var char_seq = Enumerable.Range(0, length).
                Select(k => (char)('a' + random_generator.Next(0, 25))).
                ToArray();
            return string.Concat(char_seq);
        }

        public static DataTable LightIntensityTable(IEnumerable<int> obs_range, IEnumerable<string> bands,
            IEnumerable<string> molforms)
        {
            var rand_gen = new System.Random(22556);

            // make list of gas/partial-pressure values to be used for the observations
            int molform_steps = CeilingDivide(obs_range.Count(), molforms.Count());
            var molform_concn = new List<Tuple<string, double>>();
            foreach (var form in molforms)
            {
                var gas_scale = Math.Exp(10.0 * rand_gen.NextDouble());
                var step_scale = (0.5 * rand_gen.NextDouble() + 1.25);
                var step_factor = 1.0;
                molform_concn.Add(new Tuple<string, double>(form, 0.0));
                for (int j = 1; j < molform_steps; ++j)
                {
                    double concn = Signif(gas_scale * step_factor, 4);
                    molform_concn.Add(new Tuple<string, double>(form, concn));
                    step_factor = step_factor * step_scale;
                }
            }

            var date_list = new List<DateTime>();
            var obs_id_list = new List<uint>();
            var seq_id_list = new List<uint>();
            var molform_list = new List<string>();
            var concn_list = new List<double>();
            var freq_band_list = new List<string>();
            var is_stable_list = new List<bool>();
            var intensity_list = new List<double>();
            var temperature_list = new List<double>();
            var secondary_list = new List<double>();

            var gas_iter = molform_concn.GetEnumerator();
            var prev_molform = String.Empty;
            double intensity_factor = 0.0;
            DateTime curr_date = new DateTime(2022, 1, 1, 0, 0, 0);
            double curr_temperature = 23.0 + 0.05 * rand_gen.Next(-20, 20);
            foreach (var obs_id in obs_range)
            {
                gas_iter.MoveNext();
                var curr_molform = gas_iter.Current.Item1;
                if (curr_molform != prev_molform)
                {
                    intensity_factor = 1000.0 + 300.0 * (0.5 - rand_gen.NextDouble());
                }
                prev_molform = curr_molform;
                curr_temperature += 0.02 * rand_gen.Next(-20, 20);

                uint seq_id = 1;
                foreach (var freq_band in bands)
                {
                    date_list.Add(curr_date);
                    var elapsed_seconds = (int)(24.0 + 2.0 * rand_gen.NextDouble());
                    curr_date = curr_date + new TimeSpan(0, 0, elapsed_seconds);

                    obs_id_list.Add((uint)obs_id);
                    seq_id_list.Add(seq_id++);
                    molform_list.Add(gas_iter.Current.Item1);
                    concn_list.Add(gas_iter.Current.Item2);
                    freq_band_list.Add(freq_band);
                    is_stable_list.Add(rand_gen.NextDouble() < 0.5);
                    double intensity_value = gas_iter.Current.Item2 * intensity_factor + 200.0 * (0.5 - rand_gen.NextDouble());
                    intensity_list.Add( Math.Round(intensity_value,2) );
                    temperature_list.Add(Math.Round(curr_temperature,1));
                    secondary_list.Add(0);
                }
            }

            DataTable table = new DataTable("light experiment");
            table.Columns.Add("date", typeof(DateTime));
            table.Columns.Add("obs", typeof(uint));
            table.Columns.Add("seq", typeof(uint));
            table.Columns.Add("mol formula", typeof(string));
            table.Columns.Add("concn", typeof(double));
            table.Columns.Add("color band", typeof(string));
            table.Columns.Add("stable", typeof(bool));
            table.Columns.Add("intensity", typeof(double));
            table.Columns.Add("temperature", typeof(double));
            table.Columns.Add("secondary", typeof(double));

            for (int j = 0; j < obs_id_list.Count; ++j)
            {
                DataRow new_row = table.NewRow();
                new_row["date"] = date_list[j];
                new_row["obs"] = obs_id_list[j];
                new_row["seq"] = seq_id_list[j];
                new_row["mol formula"] = molform_list[j];
                new_row["concn"] = concn_list[j];
                new_row["color band"] = freq_band_list[j];
                new_row["stable"] = is_stable_list[j];
                new_row["intensity"] = intensity_list[j];
                new_row["temperature"] = temperature_list[j];
                new_row["secondary"] = secondary_list[j];
                table.Rows.Add(new_row);
            }
            table.AcceptChanges();

            return table;
        }

        private static int CeilingDivide(int dividend, int divisor)
        {
            return (dividend + divisor - 1) / divisor;
        }

        private static double Signif(double value, int digits)
        {
            if (value == 0.0) return 0;

            var shift = -(int)Math.Ceiling(Math.Log10(Math.Abs(value))) + digits;
            var factor = Math.Pow(10.0, shift);
            int as_int = (int)(value * factor);
            return as_int / factor * Math.Sign(value);
        }
    }
}
