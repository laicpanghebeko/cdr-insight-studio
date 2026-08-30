# Politique de sécurité

## Signaler une vulnérabilité

Utilisez de préférence **Security > Report a vulnerability** sur GitHub. Si cette option n’est pas disponible, ouvrez une issue sans y joindre de document télécom, de clé API, d’identifiant personnel ou de preuve d’exploitation sensible.

Indiquez la version concernée, le scénario reproductible, l’impact attendu et une proposition de correction si vous en avez une.

## Données sensibles

- ne joignez jamais de listing réel à une issue publique ;
- reproduisez les problèmes avec le générateur synthétique `cdr_insight.demo` ;
- conservez `SERPAPI_API_KEY` dans les secrets Streamlit ou les variables d’environnement ;
- déployez derrière une authentification adaptée avant tout usage avec des données réelles.

## Versions prises en charge

La dernière version publiée sur la branche `main` reçoit les corrections de sécurité.
