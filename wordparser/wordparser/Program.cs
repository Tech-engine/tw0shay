using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

/*Author: Codeword*/
public static class globalvar
{
    public static List<string> gblDocTree = new List<string>();
    public static List<int> gblDocListNumber = new List<int>();

}
public class init_def
{
    public void init_myGlobals()
    {
        globalvar.gblDocTree.Clear();
        globalvar.gblDocListNumber.Clear();
        return;
    }
}
public class tblparam
{
    public string param;
    public tblparam(string Param = "")
    {
        param = Param;
    }
    public static int get_num(string x)
    {
        string num = x.Substring(x.IndexOf(' '));
        int j = 0;
        if (Int32.TryParse(num, out j))
            return j;
        else
            Console.WriteLine("String could not be parsed.");

        return 0;

    }
    public static void add_to_sectionnumber(int myLocation)
    {
        int myInt;
        int doc_count = globalvar.gblDocListNumber.Count;
        myInt = myLocation - 1;
        if ((myLocation > doc_count) || (doc_count == 0))
        {
            globalvar.gblDocListNumber.Add(1);
        }
        else if (doc_count == myLocation)
        {
            globalvar.gblDocListNumber[myInt] = globalvar.gblDocListNumber[myInt] + 1;
        }
        else if (doc_count > myLocation)
        {
            int y = doc_count - myLocation;
            int remove = Math.Max(0, y);
            globalvar.gblDocListNumber.RemoveRange(myLocation, remove);// equivalent to del gblDocListNumber[myLocation:]
            globalvar.gblDocListNumber[myLocation - 1] = globalvar.gblDocListNumber[myLocation - 1] + 1;
        }


    }

    public static void add_to_hierarchy(string myHeading, int myLocation)
    {
        int myInt = myLocation - 1;
        int tree_count = globalvar.gblDocTree.Count;
        if (myLocation > tree_count || tree_count == 0)
            globalvar.gblDocTree.Add(myHeading);
        else if (tree_count == myLocation)
            globalvar.gblDocTree[myInt] = myHeading;
        else if (myLocation == 1)
        {
            globalvar.gblDocTree.Clear();//del gblDocTree[:]
            globalvar.gblDocTree.Add(myHeading);
        }
        else if (tree_count > myLocation)
        {
            //int x = tree_count - myLocation;
            int y = tree_count - myLocation - 1;
            int remove = Math.Max(0, y);
            globalvar.gblDocTree.RemoveRange(myLocation - 1, remove); //equivalent to del gblDocTree[myLocation - 1:]
            globalvar.gblDocTree.Add(myHeading);
        }

    }

    public static void iter_block_items(object parent)
    {

    }
    public static int compare_two_list(List<int> curListTuple, List<int> reqStartTuple) {
        int curListTuple_length = curListTuple.Count;
        int reqStartTuple_length = reqStartTuple.Count;
        int return_val = 4;
        if (curListTuple_length < reqStartTuple_length){
            for (int i = 0; i < curListTuple_length; i++ ){
                if (curListTuple[i] > reqStartTuple[i])
                {
                    return_val = 1; // curListTuple is greater
                    break;
                }

            }
            if (return_val != 1)
            {
                return_val = 2; // reqStartTuple is greater
            }
        }
        else if (reqStartTuple_length < curListTuple_length)
        {
            for (int i = 0; i < reqStartTuple_length; i++)
            {
                if (reqStartTuple[i] > curListTuple[i])
                {
                    return_val = 2; // reqStartTuple is greater
                    break;
                }
            }
            if (return_val != 2)
            {
                return_val = 1; // curListTuple is greater
            }
        }
        else {
            for (int i = 0; i < reqStartTuple_length; i++)
            {
                if (reqStartTuple[i] > curListTuple[i])
                {
                    return_val = 2; // reqStartTuple is greater
                    break;
                }
                if (reqStartTuple[i] < curListTuple[i])
                {
                    return_val = 1; // curListTuple is greater
                    break;
                }
            }
            if (return_val != 2 && return_val != 1)
            {
                return_val = 3; // both list are equal
            }
        
        }
        return return_val;

    }
    public static string parseDocX(string mydocumentfullpath, string startSection)
    {

        init_def ini_def = new init_def();
        ini_def.init_myGlobals();   // initialization

        string myDoc = mydocumentfullpath;
        bool startSectSet = true;
        try{
            Document document = new Document(myDoc);
            //Section section = document.Sections[section_no];
            List<Node> total_nodes = new List<Node>();
            // for each section
            foreach(Section section in document.Sections){
                Body body = section.Body;
                // for each node of body of section
                
                for(Node node = body.FirstChild; node != null;node =node.NextSibling){
                    total_nodes.Add(node);
                    
                }
            }
            //Body body = section.Body;

            string prvHeader = String.Empty;
             List<string> headerLst = new List<string>{
              "Heading 1","Heading 2", "Heading 3", "Heading 4","Heading 5", "Heading 6", "Heading 7","Heading 8", "Heading 9",
              "Egemin1", "Egemin2", "Egemin3", "Egemin4","Egemin5", "Egemin6","Egemin7", "Egemin8","Egemin9", "Egemin10", "Egemin11", "Egemin12"
            };
            //bool valNext = false;
            int prvIntHeadLv = 0;
            int curHeadIntLv = 0;
            string curHeadNm = String.Empty;
            string curListNm = String.Empty;
            string myIntValName = String.Empty;
            //int myPropCnt = 0;
            Dictionary<string, string> sectionJSON = new Dictionary<string, string>();
            string paraText = String.Empty;

            string final_output = String.Empty;
            //for (Node node = total_nodes.FirstChild; node != null; node = node.NextSibling)
            foreach(Node node in total_nodes)
            {
                // Output the types of the nodes that we come across.
                if (node.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)node;

                   // Console.WriteLine(node.GetText());
                    string para_style = para.ParagraphFormat.StyleName.ToString();
                    string para_head = para.GetText();
                    if (headerLst.Contains(para_style)) { 
                        sectionJSON.Clear();
                        paraText = String.Empty;
                       
                        string curListNm_alter = String.Empty;
                        curHeadIntLv = tblparam.get_num(para_style); // eg Heading 1 gives 1
                        tblparam.add_to_hierarchy(para_head.Trim().ToLower(), curHeadIntLv);
                        if(para_head.Length >0){
                            tblparam.add_to_sectionnumber(curHeadIntLv);
                        } 
                       
                        curListNm =  string.Join(".", globalvar.gblDocListNumber.ToArray()); // 1.1.1
                        curListNm = curListNm.Trim();
                        curHeadNm = curListNm + " " + para_head;// 1.1.1 Heading1
                        string sectionHeading = String.Empty;
                        // Check if Current Section is greater than required Start
                        if (startSectSet && curListNm.Length !=0){
                            string[] words = curHeadNm.TrimStart().Split(' ');
                            sectionHeading = words[0]; //1.1.1
                            List<int> curListTuple_int = sectionHeading.Split('.').ToList().ConvertAll(s => Int32.Parse(s));
                            List<int> reqStartTuple_int = startSection.Split('.').ToList().ConvertAll(s => Int32.Parse(s));

                            if (compare_two_list(curListTuple_int, reqStartTuple_int) == 1){
                                // curListTuple is geater
                                break;
                            }
                            else if (compare_two_list(curListTuple_int, reqStartTuple_int) == 2)
                            {
                                // reqStartTuple is geater
                                continue;
                            }

                            
                        }// end of startsectset
                        if (curHeadIntLv == 1 || prvIntHeadLv == 0) {
                            prvIntHeadLv = curHeadIntLv;
                        }
                        else if( curHeadIntLv == prvIntHeadLv){
                            prvIntHeadLv = curHeadIntLv;
                            continue;
                        }

                        else if (curHeadIntLv > prvIntHeadLv)
                        {
                            prvIntHeadLv = curHeadIntLv;
                            continue;
                        }
                        else
                        {
                            prvIntHeadLv = curHeadIntLv;
                            continue;
                        }
                    }
                    else { 
                        string curParaText = para_head.Trim().ToLower();
                        paraText += curParaText;
                    }
                   
                }
                else if (node.NodeType == NodeType.Table)
                {
                    // Check if Current Section is greater than required Start
                    string sectionHeading = String.Empty;
                    if (startSectSet && curListNm.Length != 0)
                    {
                        string[] words = curHeadNm.TrimStart().Split(' ');
                        
                        sectionHeading = words[0]; //1.1.1
                        List<int> curListTuple_int = sectionHeading.Split('.').ToList().ConvertAll(s => Int32.Parse(s));
                        List<int> reqStartTuple_int = startSection.Split('.').ToList().ConvertAll(s => Int32.Parse(s));
                        //Console.WriteLine(sectionHeading + "=>" + startSection);

                        if (compare_two_list(curListTuple_int, reqStartTuple_int) == 1)
                        {
                            // curListTuple is geater
                            break;
                        }
                        else if (compare_two_list(curListTuple_int, reqStartTuple_int) == 2)
                        {
                            // reqStartTuple is geater
                            continue;
                        }

                    }// end of startsectset
                    else {
                        continue;
                    }

                    int i = 0;
                    if (curHeadNm.Length > 0)
                    {
                        string[] words = curHeadNm.TrimStart().Split(' ');
                        sectionHeading = words[0]; //1.1.1
                    }
                    else {
                        continue;
                    }

                    List<string> rowsArray = new List<string>();
                    List<string> headerArray = new List<string>();

                    /* table row loop*/
                    Table table = (Table)node;
                    foreach (Row row in table.Rows)
                    { 

                        int rowIndex = table.Rows.IndexOf(row);

                        i += 1;
                       // int myCell = 0;
                        Dictionary<string, string> JSONrow = new Dictionary<string, string>();
                        List<string> rstList = new List<string>();
                        List<string> rowStringify = new List<string>();
                        //Console.WriteLine("\tStart of Row {0}", rowIndex);

                        if (i == 1)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                headerArray.Add(cell.GetText().Trim().ToLower());
                                continue;
                            }
                        }
                        else {
                            int headerarray_len = headerArray.Count;
                            for (int j = 0; j < headerarray_len; j++)
                            {
                                //Console.WriteLine(headerArray[j]);
                                string cellText = row.Cells[j].ToString(SaveFormat.Text).Trim();
                                rowStringify.Add("\"" + headerArray[j] + "\"" + ":" + "\"" + cellText+ "\"");
                            }

                            string myStr = "{" + string.Join(",", rowStringify.ToArray()) +"}";
                            //Console.WriteLine(myStr);
                            rowsArray.Add(myStr); 
                        }
                        //Console.WriteLine("\tEnd of Row {0}", rowIndex);
                    }
                    /* table row loop stop*/
                    sectionJSON["Row"] = string.Join(",", rowsArray.ToArray());
                    break;
                }
            } // for
            foreach (KeyValuePair<string, string> entry in sectionJSON)
            {
                final_output += "{\"" + entry.Key + "\" : [" + entry.Value + "]}";
            }
            //Console.WriteLine(final_output);
            return final_output;
        }// try
        catch{
            return "Some erroe occured";
        }

    }


}
class test
{
    static void Main()
    {   
        /* note: please add references 
         1. aspose.word 
         2  system.drawing 
         before running the code
          
         Also the output is only in json string for , run the program and you can see that in the console
         */
        Console.Write(tblparam.parseDocX("example.docx", "1.1.3")); // json string output to console
        Console.ReadKey();

    }
}

