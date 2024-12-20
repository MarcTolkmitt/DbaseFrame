﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DbaseFrame
{
    /// <summary>
    /// Shortcut for 'MessageBox'.
    /// No instance no fuss.
    /// </summary>
    public class Message
    {
        /// <summary>
        /// created on: 22.10.24
        /// last edit: 05.11.24
        /// </summary>
        Version version = new Version( "1.0.2" );
        /// <summary>
        /// Shows a 'MessageBox' for convenience.
        /// </summary>
        /// <param name="text"></param>
        public static void Show( string text )
        {
            MessageBox.Show( text,
            "Message", MessageBoxButton.OK, MessageBoxImage.Error );

        }   // end: Show

        /// <summary>
        /// Query for a Yes/No and deliver the answer.
        /// </summary>
        /// <param name="text">a question for Yes/No</param>
        /// <returns>the answer</returns>
        public static bool Ask( string text )
        {
            if ( MessageBox.Show(
                    text,
                    "Query",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question ) == MessageBoxResult.Yes )
                return( true );
            else
                return ( false );

        }   // end: Ask

    }   // end: public class Message

}   // end: namespace DbaseFrame

