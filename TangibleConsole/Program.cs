using System;
using System.Data;
using System.Collections.Generic;
using Tangible;

namespace Tangible
{
    class Program
    {
        static void Main(string[] args)
        {
            if ( args.Length != 2 )
            {
                var appname = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                Console.Error.WriteLine( string.Format( "USAGE: {0} <infile> <outfile>", appname ) );
                return;
            }
            var in_filepath = args[0];
            var out_filepath = args[1];

            {
                var info = new System.IO.FileInfo(in_filepath);
                if (!info.Exists)
                {
                    Console.Error.WriteLine(string.Format("'{0}' file does not exist.", in_filepath));
                    return;
                }
            }

            var table = TsvReader.ReadTsvFile( in_filepath );
            WorkbookWriter.ExportToWorkbook(out_filepath, new List<DataTable> { table } );
        }
    }
}
