﻿using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool TryRunMacro(string path, bool save, string macroName, params object[] arguments)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)))
                return false;

            if (!System.IO.File.Exists(path))
                return false;

            bool result = false;

            Func<Workbook, bool> func = new Func<Workbook, bool>((Workbook workbook) =>
            {
                if (workbook == null)
                {
                    return false;
                }

                result = TryRunMacro(workbook.Application, out object value, macroName, arguments);

                return result && save;
            });

            Edit(path, func);

            return result;
        }

        public static bool TryRunMacro(Application application, out object result, string macroName, params object[] arguments)
        {
            result = null;

            if (application == null)
                return false;

            if (string.IsNullOrEmpty(macroName))
                return false;

            try
            {
                if(arguments == null)
                {
                    result = application.Run(macroName);
                }
                else
                {
                    switch (arguments.Length)
                    {
                        case 0:
                            result = application.Run(macroName);
                            break;

                        case 1:
                            result = application.Run(macroName, arguments[0]);
                            break;

                        case 2:
                            result = application.Run(macroName, arguments[0], arguments[1]);
                            break;

                        case 3:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2]);
                            break;

                        case 4:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3]);
                            break;

                        case 5:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4]);
                            break;

                        case 6:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]);
                            break;

                        case 7:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6]);
                            break;

                        case 8:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7]);
                            break;

                        case 9:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8]);
                            break;

                        case 10:
                            result = application.Run(macroName, arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8], arguments[10]);
                            break;
                    }
                }

                return true;
            }
            catch (Exception exception)
            {
                return false;
            }

        }
    }
}


