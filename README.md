# Delivery-automation
Automated delivery system with Google Maps API, Apps Script, and Google Sheets
Delivery Route Optimization & Data Automation System
Author: Yiheng Xiao
Platform: Google Sheets + Apps Script + Google Maps API + VLOOKUP
Project Type: Real Business Automation | MSBA Portfolio Project
Industry: Logistics / B2B Delivery

🚀 Project Overview
This project automates the daily delivery planning process for over 50+ B2B food orders by:

Standardizing customer addresses

Auto-filling delivery time windows & notes

Calculating optimized delivery routes without 3rd-party tools like Circuit.

All done within Google Sheets + Google Apps Script + Google Maps API.

🧩 Modules
1. Address Auto-Fill (Geocoding)
Uses Google Maps Geocoding API to convert company names into full standardized addresses.

Automatically updates Sheet based on customer name input.

2. Time Matching via VLOOKUP
Matches order sheet to customer sheet using VLOOKUP (or INDEX + MATCH) logic.

Auto-fills delivery time windows and special instructions.

3. Route Optimization (Custom App Script)
Uses Google Maps Directions API with optimize:true to calculate the shortest path across 50+ stops.

Supports batch processing and color-coding per delivery batch.

Fully replaces Circuit with one-click in-Sheet execution.

4. Visual Output
Dynamic Sheets with auto-updating columns, delivery notes, and sorting.

Supports export to Circuit/CSV if needed.

💻 Tech Stack
Tool	Purpose
Google Sheets	Input + Output interface
Apps Script	Automation logic
Google Maps API	Geocoding + Directions
VLOOKUP / MATCH	Time note matching
Tableau (separate)	Customer dashboard here


<img width="1673" height="376" alt="image" src="https://github.com/user-attachments/assets/5900fda8-ea00-4522-b5b8-4ba06826a26e" />
<img width="1343" height="728" alt="image" src="https://github.com/user-attachments/assets/82355c62-f437-4576-8bf1-814fdbd7446e" />

Impact
Replaced 100% manual address + time entry.

Saved ~60 mins daily on delivery planning.

Fully automated, scalable, and no paid tools needed.


