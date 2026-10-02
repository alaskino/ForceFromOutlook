using System;
using System.Collections.Generic;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace ForceFromOutlook
{
    public partial class ThisAddIn
    {
        private const string RegistryPath =
            @"Software\Alaska\ForceFromOutlook";

        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;

        // Manteniamo i riferimenti agli Explorer per evitare
        // che gli eventi COM vengano scollegati dal Garbage Collector.
        private readonly List<Outlook.Explorer> _hookedExplorers =
            new List<Outlook.Explorer>();


        // =========================================================
        // STATO ON / OFF
        // =========================================================

        internal static bool ForceFromEnabled
        {
            get
            {
                try
                {
                    using (RegistryKey key =
                        Registry.CurrentUser.CreateSubKey(RegistryPath))
                    {
                        object value = key.GetValue("Enabled", 0);

                        return Convert.ToInt32(value) == 1;
                    }
                }
                catch
                {
                    return false;
                }
            }

            set
            {
                try
                {
                    using (RegistryKey key =
                        Registry.CurrentUser.CreateSubKey(RegistryPath))
                    {
                        key.SetValue(
                            "Enabled",
                            value ? 1 : 0,
                            RegistryValueKind.DWord);
                    }
                }
                catch
                {
                    // Se il salvataggio fallisce,
                    // l'add-in continua comunque a funzionare.
                }
            }
        }


        // =========================================================
        // AVVIO ADD-IN
        // =========================================================

        private void ThisAddIn_Startup(
            object sender,
            EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;

            _explorers = Application.Explorers;
            _explorers.NewExplorer += Explorers_NewExplorer;

            // Collega gli Explorer già aperti
            for (int i = 1; i <= _explorers.Count; i++)
            {
                Outlook.Explorer explorer = _explorers[i];
                HookExplorer(explorer);
            }

            // Ultimo controllo prima dell'invio
            Application.ItemSend += Application_ItemSend;
        }


        // =========================================================
        // NUOVE FINESTRE OUTLOOK
        // =========================================================

        private void Explorers_NewExplorer(
            Outlook.Explorer Explorer)
        {
            HookExplorer(Explorer);
        }


        private void HookExplorer(
            Outlook.Explorer explorer)
        {
            if (explorer == null)
                return;

            if (_hookedExplorers.Contains(explorer))
                return;

            explorer.InlineResponse += Explorer_InlineResponse;

            _hookedExplorers.Add(explorer);
        }


        // =========================================================
        // RISPOSTA INLINE
        // =========================================================

        private void Explorer_InlineResponse(object Item)
        {
            if (!ForceFromEnabled)
                return;

            Outlook.MailItem mail = Item as Outlook.MailItem;

            if (mail != null)
            {
                ForceDefaultAccount(mail);
            }
        }


        // =========================================================
        // NUOVO INSPECTOR
        // =========================================================

        private void Inspectors_NewInspector(
            Outlook.Inspector Inspector)
        {
            if (!ForceFromEnabled)
                return;

            try
            {
                Outlook.MailItem mail =
                    Inspector.CurrentItem as Outlook.MailItem;

                if (mail != null)
                {
                    ForceDefaultAccount(mail);
                }
            }
            catch
            {
                // Nessun blocco qui:
                // ci sarà comunque il controllo finale su ItemSend.
            }
        }


        // =========================================================
        // CONTROLLO FINALE PRIMA DELL'INVIO
        // =========================================================

        private void Application_ItemSend(
            object Item,
            ref bool Cancel)
        {
            if (!ForceFromEnabled)
                return;

            Outlook.MailItem mail = Item as Outlook.MailItem;

            if (mail == null)
                return;

            Outlook.Account account = GetDefaultProfileAccount();

            if (account == null)
            {
                Cancel = true;

                System.Windows.Forms.MessageBox.Show(
                    Language.IsItalian
                        ? "Force From è attivo, ma non riesco a determinare " +
                          "l'account predefinito del profilo Outlook.\n\n" +
                          "Per sicurezza il messaggio NON è stato inviato."
                        : "Force From is enabled, but the default Outlook " +
                          "profile account could not be determined.\n\n" +
                          "For safety, the message was NOT sent.",
                    "Force From",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

                return;
            }

            try
            {
                mail.SendUsingAccount = account;
            }
            catch
            {
                Cancel = true;

                System.Windows.Forms.MessageBox.Show(
                    Language.IsItalian
                        ? "Non è stato possibile impostare l'account " +
                          "predefinito come mittente.\n\n" +
                          "Il messaggio NON è stato inviato."
                        : "The default account could not be set as the " +
                          "sender.\n\nThe message was NOT sent.",
                    "Force From",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }


        // =========================================================
        // FORZA ACCOUNT
        // =========================================================

        internal bool ForceDefaultAccount(
            Outlook.MailItem mail)
        {
            if (mail == null)
                return false;

            Outlook.Account account =
                GetDefaultProfileAccount();

            if (account == null)
                return false;

            try
            {
                mail.SendUsingAccount = account;
                return true;
            }
            catch
            {
                return false;
            }
        }


        // =========================================================
        // TROVA ACCOUNT PREDEFINITO DEL PROFILO
        // =========================================================

        internal Outlook.Account GetDefaultProfileAccount()
        {
            try
            {
                Outlook.Store defaultStore =
                    Application.Session.DefaultStore;

                Outlook.Accounts accounts =
                    Application.Session.Accounts;

                foreach (Outlook.Account account in accounts)
                {
                    try
                    {
                        Outlook.Store deliveryStore =
                            account.DeliveryStore;

                        if (deliveryStore != null &&
                            defaultStore != null &&
                            deliveryStore.StoreID == defaultStore.StoreID)
                        {
                            return account;
                        }
                    }
                    catch
                    {
                        // Alcuni tipi di account possono non avere
                        // un DeliveryStore utilizzabile.
                    }
                }
            }
            catch
            {
            }

            return null;
        }


        // =========================================================
        // ACCOUNT ATTUALMENTE FORZATO - PER TOOLTIP
        // =========================================================

        internal string GetDefaultAccountDisplayName()
        {
            try
            {
                Outlook.Account account =
                    GetDefaultProfileAccount();

                if (account == null)
                    return "";

                if (!string.IsNullOrWhiteSpace(account.SmtpAddress))
                    return account.SmtpAddress;

                return account.DisplayName;
            }
            catch
            {
                return "";
            }
        }


        // =========================================================
        // RIBBON XML
        // =========================================================

        protected override Office.IRibbonExtensibility
            CreateRibbonExtensibilityObject()
        {
            return new ForceFromRibbon();
        }


        private void ThisAddIn_Shutdown(
            object sender,
            EventArgs e)
        {
        }


        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup +=
                new EventHandler(ThisAddIn_Startup);

            this.Shutdown +=
                new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }


    // =============================================================
    // LINGUA
    // =============================================================

    internal static class Language
    {
        internal static bool IsItalian
        {
            get
            {
                try
                {
                    return
                        System.Globalization.CultureInfo
                            .CurrentUICulture
                            .TwoLetterISOLanguageName
                            .Equals(
                                "it",
                                StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}