from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithVba(DocsExamplesBase):

    def test_create_vba_project(self):

        #ExStart:CreateVbaProject
        doc = aw.Document()

        project = aw.vba.VbaProject()
        project.name = "AsposeProject"
        doc.vba_project = project

        # Create a new module and specify a macro source code.
        module = aw.vba.VbaModule()
        module.name = "AsposeModule"
        module.type = aw.vba.VbaModuleType.PROCEDURAL_MODULE
        module.source_code = "New source code"

        # Add module to the VBA project.
        doc.vba_project.modules.add(module)

        doc.save(ARTIFACTS_DIR + "WorkingWithVba.create_vba_project.docm")
        #ExEnd:CreateVbaProject

    def test_read_vba_macros(self):

        #ExStart:ReadVbaMacros
        doc = aw.Document(MY_DIR + "VBA project.docm")

        if doc.vba_project is not None:
            for module in doc.vba_project.modules:
                print(module.source_code)
        #ExEnd:ReadVbaMacros

    def test_modify_vba_macros(self):

        #ExStart:ModifyVbaMacros
        doc = aw.Document(MY_DIR + "VBA project.docm")

        project = doc.vba_project

        new_source_code = "Test change source code"
        project.modules[0].source_code = new_source_code
        #ExEnd:ModifyVbaMacros

        doc.save(ARTIFACTS_DIR + "WorkingWithVba.modify_vba_macros.docm")

    def test_clone_vba_project(self):

        #ExStart:CloneVbaProject
        doc = aw.Document(MY_DIR + "VBA project.docm")
        dest_doc = aw.Document()
        dest_doc.vba_project = doc.vba_project.clone()

        dest_doc.save(ARTIFACTS_DIR + "WorkingWithVba.clone_vba_project.docm")
        #ExEnd:CloneVbaProject

    def test_clone_vba_module(self):

        #ExStart:CloneVbaModule
        doc = aw.Document(MY_DIR + "VBA project.docm")
        dest_doc = aw.Document()
        dest_doc.vba_project = aw.vba.VbaProject()

        copy_module = doc.vba_project.modules.get_by_name("Module1").clone()
        dest_doc.vba_project.modules.add(copy_module)

        dest_doc.save(ARTIFACTS_DIR + "WorkingWithVba.clone_vba_module.docm")
        #ExEnd:CloneVbaModule

    def test_remove_broken_ref(self):

        #ExStart:RemoveReferenceFromCollectionOfReferences
        doc = aw.Document(MY_DIR + "VBA project.docm")

        # Find and remove the reference with some LibId path.
        broken_path = "brokenPath.dll"
        references = doc.vba_project.references
        for i in range(references.count - 1, -1):
            reference = doc.vba_project.references.element_at(i)
            path = get_lib_id_path(reference)
            if path == broken_path:
                references.remove_at(i)

        doc.save(ARTIFACTS_DIR + "WorkingWithVba.remove_broken_ref.docm")
        #ExEnd:RemoveReferenceFromCollectionOfReferences

    #ExStart:GetLibIdAndReferencePath
    def get_lib_id_path(self, reference):
        """Returns string representing LibId path of a specified reference."""

        if reference.type in (aw.vba.VbaReferenceType.REGISTERED, aw.vba.VbaReferenceType.ORIGINAL, aw.vba.VbaReferenceType.CONTROL):
            return self.get_lib_id_reference_path(reference.lib_id)
        if reference.type == aw.vba.VbaReferenceType.PROJECT:
            return self.get_lib_id_project_path(reference.lib_id)
        raise RuntimeError()

    @staticmethod
    def get_lib_id_reference_path(lib_id_reference: str):
        """Returns path from a specified identifier of an Automation type library.

        Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference.
        """

        if lib_id_reference is not None:
            ref_parts = lib_id_reference.split('#')
            if ref_parts.length > 3:
                return ref_parts[3]

        return ""

    @staticmethod
    def get_lib_id_project_path(lib_id_project: str):
        """Returns path from a specified identifier of an Automation type library.

        Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference.
        """

        if lib_id_project is not None:
            return lib_id_project.substring(3)

        return ""

    #ExEnd:GetLibIdAndReferencePath
