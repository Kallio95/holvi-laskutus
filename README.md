# holvi-laskutus

Laskutuksen automatisointi Holvi verkkopankkiin. Ohjelma lukee laskuaineiston Excel -tiedostosta ja lähettää sähköpostilaskut Holvia käyttäen. Excel pohjana voi käyttää Example.xlsx -tiedostoa.

## Käyttöohjeet

1. Asenna Python3.
2. Asenna python venv komennolla `python3 -m venv .` tai `python -m venv .` riippuen ympäristöstä.
3. Asenna tarvittavat kirjastot komennolla `./bin/pip install -r requirements.txt`.
4. Luo tiedosto env.json ja lisää siihen halutun Excel -tiedoston sijainti, Holvin käyttäjätunnus ja -salasana. Käytä pohjana env.json.example -tiedostoa. Huom! Excel -tiedoston tulee olla .xlsx -muodossa.
5. Aja ohjelma komennolla `./bin/python -m robot .`.
6. Selain aukeaa ja kirjautuu automaattisesti Holviin, hyväksy tarvittaessa MFA.
7. Ohjelma lukee Excel -tiedoston ja lähettää laskut Holvista.
