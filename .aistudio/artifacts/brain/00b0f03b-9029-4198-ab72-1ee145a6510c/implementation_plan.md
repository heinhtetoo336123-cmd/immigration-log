# Implementation Plan - SIA Tab Myanmar Version & Verification Upgrade

Add a full Burmese translation and digit conversion engine to all print tables, document exports (Word/HTML), and summary cards in the Statistical Investigation Analysis (SIA) tab, complete with dedicated language and digit toggle controls and verification checks.

## User Review & Critical Decisions

> [!IMPORTANT]
> The user confirmed the following product decisions during the clarification phase:

- **Language Support Mechanism**: Added interactive Language Toggle (`English` / `မြန်မာ`) across all SIA print tables, preview modals, and document exports.
- **Scope of Coverage**: Applies comprehensively to all SIA print previews (`SIA_NATIONALITY_REPORT`, `SIA_VISA_REPORT`, `SIA_FULL_REPORT`, `STILL_IN_ANALYTICS`), Word/HTML exports, and on-screen summary cards.
- **Digit Formatting**: Added a Digit Toggle (`၁, ၂, ၃` Burmese digits vs. `1, 2, 3` Western digits) that seamlessly translates all counts, ranks, percentages, and timestamps in Myanmar print mode.

---

## 1. Overview & Core Concept

- **What It Does**: Enhances the Statistical Investigation Analysis (SIA) module with a dedicated Myanmar Language and Digit system for official reporting. Allows immigration officers to toggle between official Myanmar title headers/table column headers and English versions, outputting properly translated reports for both paper prints and exported Word/HTML documents.
- **Target Audience / Persona**: Department of Immigration Officers (Myeik Second Division) generating official statistical reports for regional directors or internal auditing.
- **Key Value**: Eliminates manual translation of statistical data and ensures official Burmese terminology and digit formatting are applied consistently across all SIA print outputs.

---

## 2. User Experience & Visual Design

- **Toolbar Controls in SIA Tab**:
  - **Language Toggle Button**: `[ 🇬🇧 English | 🇲🇲 မြန်မာစာ ]` segmented toggle integrated into the SIA control toolbar.
  - **Digit Toggle Button**: `[ မြန်မာဂဏန်း (၁၂၃) | English Digits (123) ]` toggle active when Myanmar mode is selected.
  - **Verification Status Indicator**: Clean summary bar displaying verified record total and date range validation.

- **SIA Print Preview Modal & Print View (Myanmar Mode)**:
  - **Document Header**:
    - Division Title: `လူဝင်မှုကြီးကြပ်ရေးနှင့် ပြည်သူ့အင်အားဝန်ကြီးဌာန၊ မြိတ်ခရိုင်ဒုတိယဌာနခွဲ`
    - Report Header: `စင်းလုံးငှား/ပုံမှန် ခရီးစဉ်များဆိုင်ရာ ကိန်းဂဏန်းအချက်အလက် စိစစ်သုံးသပ်ချက် (SIA) အစီရင်ခံစာ`
    - Subtitle: Dynamic date range formatted in Burmese slash/dash date string (`၂၀၂၆ ခုနှစ်၊ မတ်လ ၁ ရက် မှ ၂၀၂၆ ခုနှစ်၊ မတ်လ ၁၅ ရက် အထိ`).
  - **Table 1: Nationality Inflow & Outflow (နိုင်ငံသားအလိုက် ဝင်ရောက်/ထွက်ခွာမှု စာရင်းချုပ်)**:
    - Column Headers: `စဉ်` (No.), `နိုင်ငံသား/နိုင်ငံအမည်` (Nationality), `ပြည်ဝင် (IN)` (Inbound), `ပြည်ထွက် (OUT)` (Outbound), `ကျန်ရှိ (Net Stayers)`, `စုစုပေါင်း` (Total).
    - Country Names: Standard mapping table for common country names (e.g. `THAILAND` $\rightarrow$ `ထိုင်း`, `CHINA` $\rightarrow$ `တရုတ်`, `MALAYSIA` $\rightarrow$ `မလေးရှား`, `INDIA` $\rightarrow$ `အိန္ဒိယ`, `JAPAN` $\rightarrow$ `ဂျပန်`, `SINGAPORE` $\rightarrow$ `စင်ကာပူ`, etc. with clean fallback to uppercase original string).
  - **Table 2: Visa Type Distribution (ပြည်ဝင်ဗီဇာအမျိုးအစားအလိုက် စာရင်းချုပ်)**:
    - Column Headers: `စဉ်` (No.), `ဗီဇာအမျိုးအစား` (Visa Type), `ပြည်ဝင် (IN)`, `ပြည်ထွက် (OUT)`, `စုစုပေါင်း` (Total).
    - Visa Labels: Official Burmese translations for standard visa categories (e.g. `Tourist Visa` $\rightarrow$ `ခရီးသွားဗီဇာ`, `Business Visa` $\rightarrow$ `စီးပွားရေးဗီဇာ`, `Social Visa` $\rightarrow$ `လူမှုရေးဗီဇာ`, `Special Re-entry Visa` $\rightarrow$ `အထူးပြည်ဝင်ခွင့်ဗီဇာ`, `Stay Permit` $\rightarrow$ `နေထိုင်ခွင့်`).

- **Word & HTML Exports**:
  - Updates `exportReportsToWord` and `exportReportsToHtml` to respect the active language mode and digit setting, embedding official Pyidaungsu/Unicode font styling and Burmese table structures.

---

## 3. Key Product Decisions & Trade-Offs

- **Decision 1: Centralized Translation & Digit Utility**:
  - *Chosen Approach*: Create a shared helper dictionary and digit converter (`toBurmeseDigits`, `getBurmeseCountryName`, `getBurmeseVisaTypeName`) in `src/utils/siaTranslations.ts` or inside `src/App.tsx`.
  - *Why*: Guarantees zero discrepancy between on-screen SIA preview, modal print view, raw `window.print()`, Word document generation, and HTML report files.
  - *Alternatives Considered*: Inline copy replacement in JSX was rejected to avoid code duplication across print modal vs Word/HTML generators.

- **Decision 2: Fallback Handling for Unlisted Country/Visa Strings**:
  - *Chosen Approach*: Map known common country names and visa codes to Burmese text, and safely return the clean original string if unmapped.
  - *Why*: Prevents missing data or blank fields if custom or rare country codes appear in records.

---

## 4. Technical Architecture & Data Strategy *(Technical Reference)*

```
┌────────────────────────────────────────────────────────────────────────┐
│                        SIA Tab Control Toolbar                         │
│  [Date Range]  |  [Language: EN / MM]  |  [Digits: 123 / ၁၂၃]          │
└───────────────────┬────────────────────────────────┬───────────────────┘
                    │                                │
                    ▼                                ▼
┌──────────────────────────────────────┐  ┌──────────────────────────────┐
│  On-Screen SIA Analytics Dashboard   │  │  Word / HTML Export Handlers │
│  - Summary Metric Cards (Burmese/EN) │  │  - exportReportsToWord()      │
│  - Nationality Table & Visa Table    │  │  - exportReportsToHtml()      │
└───────────────────┬──────────────────┘  └──────────────┬───────────────┘
                    │                                    │
                    ▼                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                      SIA Print Preview Modal                           │
│  - SIA_NATIONALITY_REPORT (Burmese Headers + Country Names)            │
│  - SIA_VISA_REPORT (Burmese Headers + Visa Translations)              │
│  - SIA_FULL_REPORT (Combined Burmese Statistical Analysis Report)      │
└────────────────────────────────────────────────────────────────────────┘
```

### State Mapping & Key Functions:
- `siaLanguage`: `'EN' | 'MM'` (persisted or defaulted to `'MM'`).
- `siaUseBurmeseDigits`: `boolean` (persisted or controlled by toggle).
- `toBurmeseDigits(num)`: Converts Western numbers (`0-9`) to Burmese digits (`၀-၉`).
- `translateSiaText(text, siaLanguage)`: Handles header and section title translations.
- `translateCountryName(countryName, siaLanguage)`: Maps English country names to official Burmese names.
- `translateVisaType(visaName, siaLanguage)`: Maps English visa codes/types to official Burmese names.

### Detailed Action Steps:
1. **Translation Helper Data Dictionary**:
   - Add comprehensive dictionary for country names (Thailand, China, Malaysia, India, Japan, Singapore, Vietnam, Korea, USA, UK, Russia, etc.) and visa categories.
2. **State & Control Bar Integration**:
   - Add `siaLanguage` and `siaUseBurmeseDigits` state to `StillInAnalyticsView` / `App.tsx`.
   - Add Language Segmented Toggle (`English` / `မြန်မာ`) and Digit Toggle to the top of the SIA tab control bar.
3. **Print Preview & Print DOM Updates**:
   - Update `activePrintPreview` modal sections for `'SIA_NATIONALITY_REPORT'`, `'SIA_VISA_REPORT'`, `'SIA_FULL_REPORT'`, and `'STILL_IN_ANALYTICS'`.
   - Update standalone hidden printable DOM containers to output formatted Burmese text, translated country/visa names, and Burmese numbers when Myanmar mode is enabled.
4. **Word & HTML Export Generator Updates**:
   - Pass `siaLanguage` and `siaUseBurmeseDigits` context to `exportReportsToWord` and `exportReportsToHtml`.
   - Embed font styling (`font-family: 'Pyidaungsu', 'Arial', sans-serif`) for Word/HTML files.
5. **Verification & Testing**:
   - Verify that toggling between English and Burmese dynamically updates all tables, print previews, and export outputs without layout shift or data truncation.
   - Run `compile_applet` and `lint_applet` to confirm zero compilation errors.
