function reset() {
  var service = getService();
  service.reset();
}

function getService() {
  return OAuth2.createService("GCP")
    .setTokenUrl("https://accounts.google.com/o/oauth2/token")
    .setPrivateKey(PRIVATE_KEY)
    .setIssuer(CLIENT_EMAIL)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(["https://www.googleapis.com/auth/cloud-platform"]);
}
