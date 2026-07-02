# mouseHusbandry

MATLAB tools and a GUI application for computing and tracking weekly mouse colony management tasks, including cage change scheduling, weaning dates, genotyping timelines, and available animal counts.

## Overview

Managing a transgenic mouse colony involves tracking dozens of animals across different breeding stages, genotype combinations, and weekly task schedules. This repository contains MATLAB scripts and a GUI application that automate the logic for determining which animals need attention each week — cage changes, weaning, genotyping, or experimental availability — based on birth dates, genotype records, and standard husbandry protocols.

## Contents

- `mouseColony.mlapp` — MATLAB App Designer GUI for interactive colony management; displays animals due for husbandry tasks this week and allows updating of animal records
- `mouseColony.m` / `mouseColony2.m` — Command-line versions of the colony management logic; compute weekly task lists from animal database inputs
- `availableMice.m` — Identifies animals that are available for experiments based on age, genotype, and experimental status
- `endOfRotation.m` — Computes end-of-rotation dates for animals on experimental protocols with defined time limits
- `getFolders.m` — Utility for organizing animal data directories by subject ID and date

## Requirements

- MATLAB R2018a or later
- MATLAB App Designer
