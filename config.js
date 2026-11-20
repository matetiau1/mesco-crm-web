// Konfiguracja MESco CRM Web
window.CRM_CONFIG = {
  clientId: '7277cae3-d05a-44f6-8afe-2edb6b9b6365',            // aplikacja w Entra ID
  tenantId: '4ad84be4-f918-4dfa-8006-a408554ec666',
  siteHost: 'mescocompl.sharepoint.com',
  sitePath: '/sites/CRMSystemPOC',
  scopes: ['User.Read', 'Sites.ReadWrite.All'],
  // Adres aplikacji Power Apps — pozycje menu, których nie ma w wersji web, otwierają się tam.
  // Wklej link „Odtwórz” z make.powerapps.com (np. https://apps.powerapps.com/play/e/default-.../a/...)
  powerAppsUrl: '',
  stawkaGodzinowa: 75, // € za godzinę TECS/PreSales (Bilans firm)
};
