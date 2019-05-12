using System;
using System.Windows;

// 参照に追加
using System.DirectoryServices;

namespace LDAPConnector
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            TextBaseDN.Text = "dc=stdg2wwcybbena2jljui3enboa,dc=mx,dc=internal,dc=cloudapp,dc=net";
        }

        private void ButtonConnect_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            try {
                string lserver = $"LDAP://{TextLDAPServer.Text}/{TextBaseDN.Text}";
                string domainAndUsername = $"cn={TextLDAPUser.Text},{TextBaseDN.Text}";
                string pwd = $"{TextLDAPPass.Text}";

                DirectoryEntry entry = new DirectoryEntry(lserver, domainAndUsername, pwd, AuthenticationTypes.None);
                // ここでLdap認証が入る(Exceptionが発生しなければ成功)
                object obj = entry.NativeObject;

                message = "接続成功しました";

            } catch (Exception ex) {
                // Error
                message = ex.Message;
            }
            MessageBox.Show(message);
        }

        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            try {
                string lserver = $"LDAP://{TextLDAPServer.Text}/{TextBaseDN.Text}";
                string domainAndUsername = $"cn={TextLDAPUser.Text},{TextBaseDN.Text}";
                string pwd = $"{TextLDAPPass.Text}";

                DirectoryEntry entry = new DirectoryEntry(lserver, domainAndUsername, pwd, AuthenticationTypes.None);
                //object obj = entry.NativeObject;

                var ou = entry.Children.Add("ou=testou", "OrganizationalUnit");
                ou.CommitChanges();

                message = "OUの作成に成功しました";

            } catch (Exception ex) {
                message = ex.Message;
            }
            MessageBox.Show(message);
        }

        private void ButtonAddCN_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            try {
                string lserver = $"LDAP://{TextLDAPServer.Text}/ou=testou,{TextBaseDN.Text}";
                string domainAndUsername = $"cn={TextLDAPUser.Text},{TextBaseDN.Text}";
                string pwd = $"{TextLDAPPass.Text}";

                DirectoryEntry ou = new DirectoryEntry(lserver, domainAndUsername, pwd, AuthenticationTypes.None);
                //object obj = ou.NativeObject;

                string uid = "testcn";

                var user = ou.Children.Add("cn=" + uid, "inetOrgPerson");
                user.Properties["uid"].Value = uid;
                user.Properties["sn"].Value = uid;
                user.Properties["displayName"].Value = "なまえ";
                user.Properties["homePhone"].Value = "????";
                user.Properties["homePostalAddress"].Value = "xxxxx";

                user.CommitChanges();

                message = "CNの作成に成功しました";

            } catch (Exception ex) {
                message = ex.Message;
            }
            MessageBox.Show(message);
        }

        private void ButtonRemoveOU_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            try {
                string lserver = $"LDAP://{TextLDAPServer.Text}/{TextBaseDN.Text}";
                string domainAndUsername = $"cn={TextLDAPUser.Text},{TextBaseDN.Text}";
                string pwd = $"{TextLDAPPass.Text}";

                DirectoryEntry ou = new DirectoryEntry(lserver, domainAndUsername, pwd, AuthenticationTypes.None);
                //object obj = ou.NativeObject;

                var cn = ou.Children.Find("ou=testou");

                ou.Children.Remove(cn);

                message = "OUの削除に成功しました";

            } catch (Exception ex) {
                message = ex.Message;
            }
            MessageBox.Show(message);

        }

        private void ButtonRemoveCN_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            try {
                string lserver = $"LDAP://{TextLDAPServer.Text}/ou=testou,{TextBaseDN.Text}";
                string domainAndUsername = $"cn={TextLDAPUser.Text},{TextBaseDN.Text}";
                string pwd = $"{TextLDAPPass.Text}";

                DirectoryEntry ou = new DirectoryEntry(lserver, domainAndUsername, pwd, AuthenticationTypes.None);
                //object obj = ou.NativeObject;

                var cn = ou.Children.Find("cn=testcn");

                ou.Children.Remove(cn);

                message = "CNの削除に成功しました";

            } catch (Exception ex) {
                message = ex.Message;
            }
            MessageBox.Show(message);
        }

    }
}
