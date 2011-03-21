namespace Word.W2004.Elements.TableElements
{
    /// <summary>
    /// This is returns all that crap at the top of the table, including table properties. 
    /// 
    /// It also returns the end of the table.  
    /// </summary>
    public class TableDefinition : ITableItemStrategy
    {
        public string getTop()
        {
            string top =
                 "\n	<w:tbl> "
                + "\n            <w:tblPr> "
                + "\n                <w:tblW w:w=\"0\" w:type=\"auto\"/> "
                + "\n                <w:tblBorders> "
                + "\n                    <w:top w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                    <w:left w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                    <w:bottom w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                    <w:right w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                    <w:insideH w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                    <w:insideV w:val=\"single\" w:sz=\"4\" wx:bdrwidth=\"10\" w:space=\"0\" w:color=\"000000\"/> "
                + "\n                </w:tblBorders> "
                + "\n                <w:tblLook w:val=\"00BF\"/> "
                + "\n            </w:tblPr> "
                + "\n            <w:tblGrid> "
                + "\n                <w:gridCol w:w=\"4258\"/> "
                + "\n                <w:gridCol w:w=\"4258\"/> "
                + "\n            </w:tblGrid> "
                ;
            return top;
        }

        public string getMiddle()
        {
            return null; // N/A
        }


        public string getBottom()
        {
            return "\n	</w:tbl>";
        } 
    }
}