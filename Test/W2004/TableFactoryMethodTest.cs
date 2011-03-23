using NUnit.Framework;
using Word.W2004.Elements.TableElements;

namespace Test.W2004
{
    public class TableFactoryMethodTest : Assert
    {
        [Test]
        public void testGetInstance()
        {
            TableFactoryMethod instance = TableFactoryMethod.Instance;
            Assert.NotNull(instance);
        }

        [Test]
        public void testGetTableItem()
        {
            TableFactoryMethod instance = TableFactoryMethod.Instance;

            Assert.True(TableFactoryMethod.GetTableItem(TableEle.TABLE_DEF) is TableDefinition);
            Assert.True(TableFactoryMethod.GetTableItem(TableEle.TH) is TableHeader);
            Assert.True(TableFactoryMethod.GetTableItem(TableEle.TD) is TableCol);
            Assert.True(TableFactoryMethod.GetTableItem(TableEle.TF) is TableFooter);
            Assert.True(TableFactoryMethod.GetTableItem(TableEle.TD) is TableCol);

            Assert.Null(TableFactoryMethod.GetTableItem(null));
        }
    }
}