using System.Reflection;

namespace DbaseFrame
{
    /// <summary>
    /// Paketnummer von DbaseFrame, automatisch aus der Assembly gelesen.
    /// Sie wird nur in der csproj gepflegt ( Eigenschaft Version ), AssemblyVersion und FileVersion folgen ihr.
    /// Die Klassen haben daneben weiter ihre eigenen Versionsnummern; die Paketnummer wird wie diese gepflegt:
    /// bei Änderungen erhöhen, höchstens einmal am Tag ( Vorgabe vom 27.11.26, Beginn 1.0.7 ).
    /// Erstellt:       27.11.26.
    /// Letzter Edit :  27.11.26.
    /// </summary>
    public static class PackageInfo
    {
        /// <summary>
        /// Version dieser Klasse ( nicht die Paketnummer )
        /// </summary>
        static readonly Version version = new Version( "1.0.0" );

        /// <summary>
        /// Paketnummer mit drei Stellen, z. B. 1.0.7
        /// </summary>
        public static Version Version
        {
            get
            {
                Version v = typeof( PackageInfo ).Assembly.GetName( ).Version ?? new Version( 0, 0, 0 );
                return ( new Version( v.Major, v.Minor, Math.Max( v.Build, 0 ) ) );

            }   // Ende: get

        }   // Ende: Version Version

        /// <summary>
        /// Commit, aus dem gebaut wurde ( kurz, 7 Zeichen ); das SDK hängt ihn an die InformationalVersion an.
        /// Leer, wenn keiner bekannt ist. Ungesicherte Änderungen beim Bauen zeigt er nicht.
        /// </summary>
        public static string Commit
        {
            get
            {
                string info = typeof( PackageInfo ).Assembly
                    .GetCustomAttribute<AssemblyInformationalVersionAttribute>( )?.InformationalVersion ?? "";
                int plus = info.IndexOf( '+' );
                if ( plus < 0 )
                    return ( "" );
                string commit = info.Substring( plus + 1 );
                return ( commit.Length > 7 ? commit.Substring( 0, 7 ) : commit );

            }   // Ende: get

        }   // Ende: string Commit

        /// <summary>
        /// Text für Anzeige und Protokoll, z. B. "DbaseFrame-Paket 1.0.7 ( Commit 0c24b7b )"
        /// </summary>
        public static string Text => $"DbaseFrame-Paket {Version}" + ( Commit != "" ? $" ( Commit {Commit} )" : "" );

    }   // Ende: PackageInfo

}   // Ende: namespace DbaseFrame