=================================================
Welcome to DICE! (The Digital Infrastructure Costing Model)
=================================================

.. image:: https://readthedocs.org/projects/dice-docs/badge/?version=latest
    :target: https://cdcam.readthedocs.io/en/latest/?badge=latest
    :alt: Documentation

The **Digital Infrastructure Costing Model** (``dice``) helps decision makers understand
how much annual investment is required to ensure universal broadband can plausibly be
achieved by 2030.

Citations
---------

    - TBC

Statement of Need
-----------------

Target 9c of the United Nations’ Sustainable Development Goals (SDGs) aims to ensure
that affordable universal broadband connectivity is available to all citizens, given
the widely acknowledged economic development benefits.

However, there are few tools that can guide decision makers to understand the required
investment needs to achieve SDG Target 9c. There are even fewer examples of tools which
can help support cross-country comparisons using a consistent methodological approach.

Additionally, using the DICE model allows users to take advantage of either the baseline
model parameters, or user-defined parameter values, enabling a left of interactivity
which helps to boost analyst understanding.

Applications
------------

There are many applications for the DICE model. The two main applications include:

- Exploring the cost of delivering different quality of service levels to users, in terms
of the per user broadband speed achievable (e.g., 10 Mbps per user).
- Exploring the cost of using different broadband infrastructure strategies to achieve
SDG Target 9c (e.g., using 4G).


Setup and configuration
-----------------------

Non-technical users can download the excel spreadsheet present in this repository titled
`Oughton et al. (2022) DICE.xlsx`.

Or, technical users can clone this repository and run the following to extract population
information (you will need to download the WorldPop 2020 population layer):

    python scripts/pop.py

And then run the following to build the DICE spreadsheet model:

    python scripts/build.py


Background and funding
----------------------

The development of this software has been supported by the International Monetary Fund's
(IMF's) Digital Infrastructure Costing Estimator (DICE) project.


Contents
--------

.. toctree::
   :maxdepth: 2

   Getting Started <getting-started>

.. toctree::
   :maxdepth: 1

   License <license>
   Authors <authors>
