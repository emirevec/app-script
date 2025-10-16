function getCurrentUserEmailID() {
  const currentUser = Session.getActiveUser().getEmail();
  const username = currentUser.includes('@') 
    ? currentUser.split('@')[0].toLowerCase() 
    : currentUser.toLowerCase();
  
  return XLXS_USERS_NAMES[username] || currentUser;
}