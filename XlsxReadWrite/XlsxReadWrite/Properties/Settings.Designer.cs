﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace XlsxReadWrite.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("XlsxReadWriteTable")]
        public string DataTableName {
            get {
                return ((string)(this["DataTableName"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("XlsxReadWrite")]
        public string DataTableNamespace {
            get {
                return ((string)(this["DataTableNamespace"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("TEMP")]
        public string TemporaryDirectory {
            get {
                return ((string)(this["TemporaryDirectory"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("StateRoster.xlsx")]
        public string InputFileName {
            get {
                return ((string)(this["InputFileName"]));
            }
            set {
                this["InputFileName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("newAMCRoster.xlsx")]
        public string OutputFileName {
            get {
                return ((string)(this["OutputFileName"]));
            }
            set {
                this["OutputFileName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("AMCRoster.xlsx")]
        public string InputFileName2 {
            get {
                return ((string)(this["InputFileName2"]));
            }
            set {
                this["InputFileName2"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1,0,2,3,4,0,0,7,0,8,5,6")]
        public string Input1Columns {
            get {
                return ((string)(this["Input1Columns"]));
            }
            set {
                this["Input1Columns"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1,2,3,4,5,6,7,8,9,10,11,12,13,14")]
        public string Input2Columns {
            get {
                return ((string)(this["Input2Columns"]));
            }
            set {
                this["Input2Columns"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("AEC,Adelaide Showground,Adelaide Oval,Adelaide (M/C),Woodville,Pt Adelaide")]
        public string IncludeMatchArray_str {
            get {
                return ((string)(this["IncludeMatchArray_str"]));
            }
            set {
                this["IncludeMatchArray_str"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Metro North East Region,Metro South Region,Mid North & Yorke Peninsula Region,Riv" +
            "erland & South East Region,South Coast & Adelaide Hills Region,Western Region")]
        public string RemoveMatchArray_str {
            get {
                return ((string)(this["RemoveMatchArray_str"]));
            }
            set {
                this["RemoveMatchArray_str"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("6,6,6,8,8,8")]
        public string IncludeMatchArray_col {
            get {
                return ((string)(this["IncludeMatchArray_col"]));
            }
            set {
                this["IncludeMatchArray_col"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("7,7,7,7,7,7")]
        public string RemoveMatchArray_col {
            get {
                return ((string)(this["RemoveMatchArray_col"]));
            }
            set {
                this["RemoveMatchArray_col"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("false,false,false,false,false,false")]
        public string IncludeMatchArray_exact {
            get {
                return ((string)(this["IncludeMatchArray_exact"]));
            }
            set {
                this["IncludeMatchArray_exact"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("true,true,true,true,true,true")]
        public string RemoveMatchArray_exact {
            get {
                return ((string)(this["RemoveMatchArray_exact"]));
            }
            set {
                this["RemoveMatchArray_exact"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool FilterHidden {
            get {
                return ((bool)(this["FilterHidden"]));
            }
            set {
                this["FilterHidden"] = value;
            }
        }
    }
}
