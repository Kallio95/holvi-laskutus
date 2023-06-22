# holvi-laskutus
Laskutuksen automatisointi Holvi verkkopankkiin. Ohjelma lukee laskuaineiston Excel -tiedostosta ja lähettää sähköpostilaskut Holvia käyttäen.

## Käyttöohjeet
1. Asenna Python3.
2. Asenna tarvittavat kirjastot komennolla `pip install -r requirements.txt` tai `pip3 install -r requirements.txt` riippuen ympäristöstä.
3. Luo tiedosto env.json ja lisää siihen halutun Excel -tiedoston sijainti, Holvin käyttäjätunnus ja -salasana. Käytä pohjana env.json.example -tiedostoa. Huom! Excel -tiedoston tulee olla .xlsx -muodossa.
4. Aja ohjelma komennolla `python -m robot .` tai `python3 -m robot .` riippuen ympäristöstä.
5. Selain aukeaa ja kirjautuu automaattisesti Holviin, hyväksy tarvittaessa MFA.
6. Ohjelma lukee Excel -tiedoston ja lähettää laskut Holvista.