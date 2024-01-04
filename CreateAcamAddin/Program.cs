using System.IO;
using System.Text;

namespace CreateAcamAddin
{
    class Program
    {
        /*
         * $1 ' Name of .Net DLL to load, no path or extension.
        [file name]

        $2 ' 1 to always load, 0 to show in Utils | Add-Ins dialog box
        0

        $3 ' NoDefaultLoad, = 1 don't load if user registry entry not set. (Only used if $2 = 0)
        0

        $4 ' 1 if this is an Extra Cost Option
        0

        $10 ' Load for Mill, 1 = yes, 0 = no
        1

        $20 ' Load for Router, 1 = yes, 0 = no
        1

        $30 ' Load for Stone, 1 = yes, 0 = no
        1

        $40 ' Load for Lathe, 1 = yes, 0 = no
        0

        $50 ' Load for Profiling, 1 = yes, 0 = no
        0

        $60 ' Load for Wire, 1 = yes, 0 = no
         */
        static void Main ( string[] args )
        {
            if ( args.Length != 2 )
            {
                return;
            }

            string strOutputName = getOutputName(args);
            string strOutputDir = getDirectory(args);
            string strOutputFile = $"{strOutputDir}{strOutputName}.acamaddin";

            StringBuilder objBuilder = new StringBuilder();
            objBuilder.AppendLine ( "$1 ' Name of .Net DLL to load, no path or extension." );
            objBuilder.AppendLine ( strOutputName );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$2 ' 1 to always load, 0 to show in Utils | Add-Ins dialog box" );
            objBuilder.AppendLine ( "0" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$3 ' NoDefaultLoad, = 1 don't load if user registry entry not set. (Only used if $2 = 0)" );
            objBuilder.AppendLine ( "0" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$4 ' 1 if this is an Extra Cost Option");
            objBuilder.AppendLine ( "0" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$10 ' Load for Mill, 1 = yes, 0 = no");
            objBuilder.AppendLine ( "1");
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$20 ' Load for Router, 1 = yes, 0 = no:" );
            objBuilder.AppendLine ( "1" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$30 ' Load for Stone, 1 = yes, 0 = no" );
            objBuilder.AppendLine ( "1" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$40 ' Load for Lathe, 1 = yes, 0 = no" );
            objBuilder.AppendLine ( "0" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$50 ' Load for Profiling, 1 = yes, 0 = no" );
            objBuilder.AppendLine ( "0" );
            objBuilder.AppendLine ( );
            objBuilder.AppendLine ( "$60 ' Load for Wire, 1 = yes, 0 = no" );
            objBuilder.AppendLine ( "0");

            File.WriteAllText (strOutputFile, objBuilder.ToString ( ), Encoding.ASCII );
        }

        static string getDirectory ( string[] args )
        {
            foreach ( var arg in args )
            {
                if ( Directory.Exists ( arg ) )
                {
                    return arg;
                }
            }

            return string.Empty;
        }

        static string getOutputName ( string[] args )
        {
            foreach ( var arg in args )
            {
                if ( Directory.Exists ( arg ) )
                {
                    continue;
                }

                return arg;
            }

            return string.Empty;
        }
    }
}
