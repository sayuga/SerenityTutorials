
namespace myProject.Default.Forms
{
    using Serenity;
    using Serenity.ComponentModel;
    using Serenity.Data;
    using System;
    using System.ComponentModel;
    using System.Collections.Generic;
    using System.IO;

    [FormScript("Default.ExcelImport")]
    public class ExcelImportForm
    {
        [FileUploadEditor, Required]
        public String FileName { get; set; }
    }

}
