
# LUP - Lean Upgrade Plan

## Description

LUP est un outil de gestion et de suivi d'actions développé sous Microsoft Excel, VBA et Python dans Excel.

Le projet a été conçu pour centraliser les actions, automatiser la création des fiches de suivi, simplifier le pilotage opérationnel et produire des indicateurs de performance (KPI).

Cette première version repose sur Excel et VBA, avec une évolution prévue vers un dashboard web basé sur Python, Streamlit et PostgreSQL.


# Objectifs du projet

- Centraliser le suivi des actions
- Faciliter le pilotage des activités
- Réduire les tâches manuelles
- Améliorer la visibilité sur l'avancement des sujets
- Générer automatiquement des KPI
- Préparer une future migration vers une architecture Web



# Fonctionnalités

## Gestion des actions

- Création d'actions
- Modification d'actions
- Suivi du statut
- Gestion des acteurs
- Gestion des priorités
- Gestion des criticités

## Génération automatique des fiches

Chaque action possède une fiche individuelle générée automatiquement.

Fonctionnalités :

- Création automatique des onglets
- Mise à jour des fiches existantes
- Suppression des fiches obsolètes
- Synchronisation Sommaire ↔ Fiches
- Génération des hyperliens
- Navigation entre les feuilles


## Navigation

### Depuis le Sommaire

- Accès à la fiche détaillée via un lien "Voir le détail"

### Depuis la fiche

- Retour automatique vers le Sommaire


## Référentiels dynamiques

Les listes déroulantes sont maintenues depuis la feuille **Tutoriel**.

Référentiels disponibles :

- Type d'action
- Criticité
- Priorité
- Statut action

L'ajout de nouvelles valeurs ne nécessite aucune modification du VBA.


# Architecture du classeur


# Technologies utilisées

## Excel

- Validation de données
- Hyperliens
- Formules
- Mise en forme conditionnelle

## VBA

- Génération automatique des fiches
- Synchronisation des données
- Navigation
- Calcul des KPI
- Gestion des référentiels

## Python dans Excel

Bibliothèques utilisées :

```python
pandas
matplotlib
```

Utilisation :

- Analyse de données
- Agrégation
- Répartition des charges
- Préparation du dashboard



# KPI disponibles

## Actions ouvertes

Nombre d'actions dont le statut est différent de :

Done
Cancelled


Objectif :

Identifier la charge de travail restante.

-
## Actions fermées

Nombre d'actions dont le statut est :

```text
Done
Cancelled
```

Objectif :

Mesurer l'avancement global du portefeuille d'actions.


## Actions non modifiées depuis plus de 30 jours

Basé sur :


Date dernière modification


Objectif :

Identifier les sujets qui ne sont plus suivis.



## Actions créées depuis plus de 30 jours

Basé sur :

Date de création


Objectif :

Identifier les sujets anciens encore présents dans le portefeuille.



## Répartition des statuts

Suivi des statuts :

En cours
Done
Cancelled
Analysing in progress
En standby
Implementing in progress
Validating in progress




## Types d'action non modifiés depuis plus de 30 jours

Analyse des catégories de sujets qui ne sont plus mises à jour :

Evolution
ER
Incident
Méthode
Migration
Norme
Processus
SR


# KPI Python

## Nombre de sujets P1 par acteur

Permet d'identifier les collaborateurs portant les sujets les plus critiques.

Exemple :


Mathieu      18
Naïm         11
Seydou        3


## Nombre total de sujets par acteur
## Visualisations

Graphiques générés :

- Nombre de sujets par acteur
- Nombre de sujets P1 par acteur
- Répartition des priorités



# Tests réalisés

## Génération des fiches

- Création automatique
- Mise à jour
- Suppression des fiches inutiles

## Navigation

- Sommaire → Fiche
- Fiche → Sommaire




## Référentiels

- Ajout dynamique de valeurs
- Mise à jour des listes déroulantes


## KPI

- Actions ouvertes
- Actions fermées
- Actions créées > 30 jours
- Actions non modifiées > 30 jours
- Répartition des statuts


# Structure du dépôt


└── LUP_Demo.xlsm



# Évolutions prévues

## V2

### Pilotage des releases

- Nombre de sujets par release
- Historique des livraisons
- Répartition des sujets par sprint
- Timeline des releases

### Charge et performance

- Nombre de sujets P1 par acteur
- Répartition des tâches
- Charge par pilote
- Analyse des priorités



## Évolution du projet
Ce projet constitue la première version du produit.
La feuille de route prévoit la migration vers une architecture Web basée sur :
- Python
- Streamlit
- PostgreSQL

afin de dissocier :
- la saisie des données (Excel)
- l'analyse et le pilotage (Dashboard)

# Auteur

Projet réalisé dans une démarche d'amélioration continue et d'automatisation de processus métier.

Objectif :

Transformer un suivi manuel sous Excel en une solution de pilotage automatisée, maintenable et évolutive vers un futur système décisionnel basé sur une architecture Web.
