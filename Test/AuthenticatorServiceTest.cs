using NUnit.Framework;
using Word.Security;

namespace Test
{
    public class AuthenticatorServiceTest : Assert
    {
        [Test]
        public void AuthenticatorServiceTestTest()
        {
            AuthenticatorService @as = new AuthenticatorService();
            Assert.True(AuthenticatorService.Authenticate());
        }
    }
}