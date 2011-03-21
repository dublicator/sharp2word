namespace Word.W2004.Elements.TableElements
{
    /// <summary>
    /// Here is the logic to decide which instance create and return 
    /// </summary>
    public class TableFactoryMethod
    {
        private static TableFactoryMethod instance;

        private TableFactoryMethod()
        {
        }

        public static TableFactoryMethod getInstance()
        {
            if (instance == null)
            {
                instance = new TableFactoryMethod();
            }
            return instance;
        }

        public ITableItemStrategy getTableItem(TableEle tableEle)
        {
            if (tableEle == null)
            {
                return null;
            }

            return getTableEle(tableEle);
        }

        private ITableItemStrategy getTableEle(TableEle tableEle)
        {
            if (tableEle.getValue().Equals("tableDef"))
            {
                return new TableDefinition();
            }
            else if (tableEle.getValue().Equals("th"))
            {
                return new TableHeader();
            }
            else if (tableEle.getValue().Equals("td"))
            {
                return new TableCol();
            }
            else
            { //if (tableEle.getValue().equals("tf")) {
                return new TableFooter();
            }
        } 
    }
}