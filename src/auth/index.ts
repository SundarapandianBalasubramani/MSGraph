export const config = {
    appId: '00000000-0000-0000-0000-000000000000',
    authority: 'https://login.microsoftonline.com/00000000-0000-0000-0000-000000000000',
    redirectUri: 'http://localhost:8080',
    scopes: ['User.Read', 'People.Read', 'User.Read.All', 'Calendars.Read',
     'GroupMember.Read.All', 'User.ReadBasic.All', 'People.Read.All',
      'Presence.Read.All', 'Sites.Read.All', 'Mail.ReadBasic', 'mailboxsettings.read',
      'Team.ReadBasic.All', 'TeamSettings.Read.All',  
      'Directory.Read.All','TeamMember.Read.All']
};