import unittest
import os
import sys

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class DocumentProtection(DocsExamplesBase):

    def test_protect(self):

        #ExStart:ProtectDocument
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.protect(aw.ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password")
        #ExEnd:ProtectDocument

    def test_unprotect(self):

        #ExStart:UnprotectDocument
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.unprotect()
        #ExEnd:UnprotectDocument

    def test_get_protection_type(self):

        #ExStart:GetProtectionType
        doc = aw.Document(MY_DIR + "Document.docx")
        protection_type = doc.protection_type
        #ExEnd:GetProtectionType

    def test_read_only_protection(self):

        #ExStart:ReadOnlyProtection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Open document as read-only")

        # Enter a password that's up to 15 characters long.
        doc.write_protection.set_password("MyPassword")

        # Make the document as read-only.
        doc.write_protection.read_only_recommended = True

        # Apply write protection as read-only.
        doc.protect(aw.ProtectionType.READ_ONLY)
        doc.save(ARTIFACTS_DIR + "DocumentProtection.ReadOnlyProtection.docx")
        #ExEnd:ReadOnlyProtection

    def test_remove_read_only_restriction(self):

        #ExStart:RemoveReadOnlyRestriction
        doc = aw.Document()

        # Enter a password that's up to 15 characters long.
        doc.write_protection.set_password("MyPassword")

        # Remove the read-only option.
        doc.write_protection.read_only_recommended = False

        # Apply write protection without any protection.
        doc.protect(aw.ProtectionType.NO_PROTECTION)
        doc.save(ARTIFACTS_DIR + "DocumentProtection.RemoveReadOnlyRestriction.docx")
        #ExEnd:RemoveReadOnlyRestriction

    def test_password_protection(self):

        #ExStart:PasswordProtection
        doc = aw.Document()

        # Apply document protection.
        doc.protect(aw.ProtectionType.NO_PROTECTION, "password")

        doc.save(ARTIFACTS_DIR + "DocumentProtection.PasswordProtection.docx")
        #ExEnd:PasswordProtection

    def test_allow_only_form_fields_protect(self):

        #ExStart:AllowOnlyFormFieldsProtect
        # Insert two sections with some text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Text added to a document.")

        # A document protection only works when document protection is turned and only editing in form fields is allowed.
        doc.protect(aw.ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password")

        # Save the protected document.
        doc.save(ARTIFACTS_DIR + "DocumentProtection.AllowOnlyFormFieldsProtect.docx")
        #ExEnd:AllowOnlyFormFieldsProtect

    def test_remove_document_protection(self):

        #ExStart:RemoveDocumentProtection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Text added to a document.")

        # Documents can have protection removed either with no password, or with the correct password.
        doc.unprotect()
        doc.protect(aw.ProtectionType.READ_ONLY, "newPassword")
        doc.unprotect("newPassword")

        doc.save(ARTIFACTS_DIR + "DocumentProtection.RemoveDocumentProtection.docx")
        #ExEnd:RemoveDocumentProtection

    def test_unrestricted_editable_regions(self):

        #ExStart:UnrestrictedEditableRegions
        # Upload a document and make it as read-only.
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)

        doc.protect(aw.ProtectionType.READ_ONLY, "MyPassword")

        builder.writeln("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.")

        # Start an editable range.
        ed_range_start = builder.start_editable_range()
        # An EditableRange object is created for the EditableRangeStart that we just made.
        editable_range = ed_range_start.editable_range

        # Put something inside the editable range.
        builder.writeln("Paragraph inside first editable range")

        # An editable range is well-formed if it has a start and an end.
        ed_range_end = builder.end_editable_range()

        builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.")

        doc.save(ARTIFACTS_DIR + "DocumentProtection.UnrestrictedEditableRegions.docx")
        #ExEnd:UnrestrictedEditableRegions

    def test_unrestricted_section(self):

        #ExStart:UnrestrictedSection
        # Insert two sections with some text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Section 1. Unprotected.")
        builder.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
        builder.writeln("Section 2. Protected.")

        # Section protection only works when document protection is turned and only editing in form fields is allowed.
        doc.protect(aw.ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password")

        # By default, all sections are protected, but we can selectively turn protection off.
        doc.sections[0].protected_for_forms = False
        doc.save(ARTIFACTS_DIR + "DocumentProtection.UnrestrictedSection.docx")

        doc = aw.Document(ARTIFACTS_DIR + "DocumentProtection.UnrestrictedSection.docx")
        self.assertFalse(doc.sections[0].protected_for_forms)
        self.assertTrue(doc.sections[1].protected_for_forms)
        #ExEnd:UnrestrictedSection
