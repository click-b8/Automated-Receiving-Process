# templates/

Put your company's PRF template here (QF 7.4.3-02, the version with native Excel checkboxes and the letterhead logo) as:

    PRF template CONTROLS.xlsx

`prf_helpers.FALLBACK_LOGO_TEMPLATE` points at this file. `export_and_stamp_pdf()` uses its embedded logo (`xl/media/image1.png`) when a PRF being exported has lost its own. If neither has a logo, that part is BLOCKED (`LogoMissingError`) instead of exporting an unbranded record.

The template is a controlled company document and is not part of this repository.
