    /**
     * 
     * @param args {args.NameInput:string,args.NameOutput:string}
     */
    var WordConvert = function (args) {
        
        try {
        // ************************************************************************ 
        // PRIVATE VARIABLES AND FUNCTIONS         
        // ***********************************************************************
            var FSO = new ActiveXObject("Scripting.FileSystemObject");
            
            
            /**
             * 
             * @param FileName
             * @returns path or null
             */
            var GetFullPath = function (FileName) {
                var path = null;
                if (FSO.FileExists(FileName)) path = FSO.GetAbsolutePathName(FileName);
                
                if (!path) throw new Error("ѕуть не найден:" + FileName);
                return path;
            };
            /**
             * 
             * @param foldername
             * @returns fullpath
             */
            var GetOutFolderPath = function (foldername)
            {
                var path = null;
                var treeArray = foldername.split(/([\\/])/)
                
                
                var namefile = treeArray[treeArray.length - 1];
                
                treeArray.splice(treeArray.length - 1, 1)
                if (treeArray[0] == "..") { treeArray[0] = FSO.GetAbsolutePathName(".\\"); }
                var pathtofolder = treeArray.join('\\')
                
                if (FSO.FolderExists(pathtofolder)) path = FSO.GetAbsolutePathName(pathtofolder);

                if (path == null) throw new Error("ѕуть не найден:" + path);
                
                return path + "\\" + namefile;
            }

            var NameInput = "";
            var NameOutput = "";
            var PathInput = "";
            var PathOutput = "";
            
            if (args!=null) {
                if (args.NameInput!=null) { NameInput = args.NameInput; PathInput = GetFullPath(NameInput); }
                if (args.NameOutput!=null) { NameOutput = args.NameOutput; PathOutput = GetOutFolderPath(NameOutput); }

            }

            
            
            
            
            
            /**
             * 
             * @param file
             */
            var DeleteFile =function(file)
            {
                var path_f = GetFullPath(file);
                FSO.DeleteFile(path_f);
            }
            /**
             * 
             * @returns HTML
             */
            var GetHTMLfromFile = function () {
                var myModelessDialog = showModelessDialog(PathOutput, window, "status:false; dialogWidth:400px;dialogHeight:400px;help:no;status:no;center:yes;resizable:yes;minimize:yes;maximize:yes;scroll:no;");
                var modelessBody = myModelessDialog.document.body;
                var HTML = myModelessDialog.document.getElementsByTagName('html')[0].innerHTML;
                myModelessDialog.close();
                return HTML;
            }
            // ************************************************************************ 
            // PRIVILEGED METHODS 
            // MAY BE INVOKED PUBLICLY AND MAY ACCESS PRIVATE ITEMS        
            // ************************************************************************ 
            /**
             *          
             * @param outdocumentname
             * @returns
             */            
            this.SaveAsHTML = function (outdocumentname) {
                if (outdocumentname!=null) {
                    NameOutput = outdocumentname;
                    PathOutput = GetOutFolderPath(outdocumentname);
                }
                if (PathOutput == "")
                {
                    NameOutput = ".\\Default.html";
                    PathOutput = GetOutFolderPath(NameOutput);
                }
                
                var wrdApp = new ActiveXObject("Word.Application");
                var wrdDoc = wrdApp.Documents.Open(PathInput);
                wrdApp.ActiveDocument.SaveAs(PathOutput, 8); // 8 represents the html fileformat
                wrdDoc.Close(0);
                wrdApp.Quit();
            }            
            /**
             * 
             * @param inputname
             */
            this.Open = function (inputname) {
                if (!inputname) {
                    NameInput = inputname;
                    PathInput = GetFullPath(inputname);
                }
            }
            /**
             * 
             * @returns HTML
             */
            this.GetHTML = function ()
            {
                
                return GetHTMLfromFile();
            }
            /**
             * @fires SaveAsHTML  
             * @returns HTML
             */
            this.GenerateHTML = function()
            {
                if (!PathInput) throw new Error("¬ходной файл не установлен");
                this.SaveAsHTML();                                
            }
            /**
             * 
             * @returns path to Default.html
             */
            this.GetSrcHTML = function ()
            {
                return PathOutput;
            }
        }
        catch (e)
        {
            alert(e.message + "\nLine:" + e.lineNumber);
        }
    }