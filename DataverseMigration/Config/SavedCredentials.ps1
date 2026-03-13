# Opcjonalnie: pelna konfiguracja w PowerShell (URL + login + haslo per zrodlo/cel).
# Prostsza metoda: uzyj pliku Config\LoginHaslo.txt (tylko Login= i Haslo=), a URL wpisuj w aplikacji.
# NIE commituj tego pliku z haslami do repozytorium.

# --- ZRODLO ---
$SourceUrl      = ''   # np. https://mojaorg.crm4.dynamics.com
$SourceLogin    = ''   # np. user@firma.com
$SourcePassword = ''   # haslo do konta

# --- CEL ---
$TargetUrl      = ''
$TargetLogin    = ''
$TargetPassword = ''

# Opcjonalnie: wlasny AppId i RedirectUri (domyslnie uzywane sa standardowe dla Dynamics)
# $OAuthAppId      = '51f81489-12ee-4a9e-aaae-a2591f45987d'
# $OAuthRedirectUri = 'app://58145B91-0C36-4500-8554-080854F2AC97'
