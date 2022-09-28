---
title: test portfolio
icon: fas fa-link
image: /imgs/data-model/data-model-pivot.png
order: 5
show: true
---

## excel

{% assign sites = site.categories['Excel'] | sort: 'date' | reverse %}
{% include cards.html references = sites %}

## word

{% assign publications = site.categories['Word'] | sort: 'date' | reverse %}
{% include cards.html references = publications %}
