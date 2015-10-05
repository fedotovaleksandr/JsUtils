// ************************************************************************ 
// Work with File system and files     
// ************************************************************************



var FileSystem = function ()
{
    // ************************************************************************ 
    //_private    
    // ************************************************************************
    
    

    // ************************************************************************ 
    //public 
    // ************************************************************************
    
    
    this._FSO = new ActiveXObject("Scripting.FileSystemObject");
}

/**
 * 
 * @param {String} folderpath
 * @returns {Boolean} 
 */
FileSystem.prototype.FolderExsists = function (folderpath) {
    
    return this._FSO.FolderExists(folderpath)
}
/**
 * 
 * @param {String} path
* @returns {Boolean}
 */
FileSystem.prototype.FileExsists = function (path) {
   
    return this._FSO.FileExists(path)
}
/**
 * 
 * @param {String} path
* @returns {Boolean}
*  @throws {Error}
 */
FileSystem.prototype.FolderCreate = function (path) {
    var answer =false;
    try {
        this._FSO.CreateFolder(path);
        answer = true;
    }
    catch (e) {

        throw new Error("Error Create Folder");

    }
    finally {
        return answer;
    }
}
/**
 * 
 * @param {String} source
 * @param {String} destination
 * @param {Boolean} overwrite
*  @returns {String Boolean} dest
* @throws {Error}
 */
FileSystem.prototype.CopyFile = function (source, destination, overwrite) {
    var answer = false;
    try {
        //check
        if (!this.FileExsists(source)) {throw new Error("File "+source+" Not Exist")};
        
        if (destination[destination.length - 1] == "/") destination += "default";
        if (this.FileExsists(destination)) {
            
            var Words_Array = destination.split(/\./);
            var type=Words_Array[Words_Array.length - 1]
            destination = "\." + Words_Array.slice(0, Words_Array.length - 1).join("") + this.CreateTimeFileName() + "\." + type;
            

        };
        //copy
        this._FSO.CopyFile(source, destination, overwrite);
        answer = true;
    }
    catch (e) {
        alert(e.message);
        throw new Error(e.message);

    }
    finally {
        return { dest: destination, answer: answer };
    }
}
/**
 * 
 * @param source
 */
FileSystem.prototype.GetFolder = function (source)
{
    try {
        if (this.FolderExsists(source)) {
            return this._FSO.GetFolder(source);
        }
    } catch (e) {
        alert(e.message);
        throw new Error(e.message);
    }
}
/**
 * 
 * @param source
* @description delete and check file from source
 */
FileSystem.prototype.DeleteFile = function (source)
{
    try{
        if (this.FileExsists(source))
        {
            this._FSO.DeleteFile(source);
        }
    } catch (e) {
        alert(e.message);
        throw new Error(e.message);
    }
    
}
/**
 * 
 * @param source
 * @param recursive
 */
FileSystem.prototype.DeleteFolder = function (source, recursive)
{
    try {
        if (this.FolderExsists(source)) {
            this._FSO.DeleteFolder(source);
        }
    } catch (e) {
        alert(e.message);
        throw new Error(e.message);
    }
}


/**
 * @returns {String}
* @description format YYYY_MM_DD_timeHH_SS
 */
FileSystem.prototype.CreateTimeFileName = function () {

    var time = new Date();
    var time_year = time.getFullYear();
    var time_mounth = time.getMonth() + 1;
    var time_day = time.getDate();
    var time_min = time.getMinutes();
    var time_hours = time.getHours();
    var time_sec = time.getSeconds();
    var time_wr = time_year + "" + time_mounth + "" + time_day + "time"
    time_wr += ((time_hours < 10) ? "0" : "") + time_hours;
    time_wr += "";
    time_wr += ((time_min < 10) ? "0" : "") + time_min;
    time_wr += "" + time_sec;


    return time_wr;

}

