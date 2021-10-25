import unittest
import os
import sys
from datetime import datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithComments(docs_base.DocsExamplesBase):

    def test_add_comments(self):

        #ExStart:AddComments
        #ExStart:CreateSimpleDocumentUsingDocumentBuilder
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Some text is added.")
        #ExEnd:CreateSimpleDocumentUsingDocumentBuilder

        comment = aw.Comment(doc, "Awais Hafeez", "AH", datetime.today())

        builder.current_paragraph.append_child(comment)

        comment.paragraphs.add(aw.Paragraph(doc))
        comment.first_paragraph.runs.add(aw.Run(doc, "Comment text."))

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.add_comments.docx")
        #ExEnd:AddComments


    def test_anchor_comment(self):

        #ExStart:AnchorComment
        doc = aw.Document()

        para1 = aw.Paragraph(doc)
        run1 = aw.Run(doc, "Some ")
        run2 = aw.Run(doc, "text ")
        para1.append_child(run1)
        para1.append_child(run2)
        doc.first_section.body.append_child(para1)

        para2 = aw.Paragraph(doc)
        run3 = aw.Run(doc, "is ")
        run4 = aw.Run(doc, "added ")
        para2.append_child(run3)
        para2.append_child(run4)
        doc.first_section.body.append_child(para2)

        comment = aw.Comment(doc, "Awais Hafeez", "AH", datetime.today())
        comment.paragraphs.add(aw.Paragraph(doc))
        comment.first_paragraph.runs.add(aw.Run(doc, "Comment text."))

        comment_range_start = aw.CommentRangeStart(doc, comment.id)
        comment_range_end = aw.CommentRangeEnd(doc, comment.id)

        run1.parent_node.insert_after(comment_range_start, run1)
        run3.parent_node.insert_after(comment_range_end, run3)
        comment_range_end.parent_node.insert_after(comment, comment_range_end)

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.anchor_comment.doc")
        #ExEnd:AnchorComment


    def test_add_remove_comment_reply(self):

        #ExStart:AddRemoveCommentReply
        doc = aw.Document(docs_base.my_dir + "Comments.docx")

        comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()
        comment.remove_reply(comment.replies[0])

        comment.add_reply("John Doe", "JD", datetime(2017, 9, 25, 12, 15, 0), "New reply")

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.add_remove_comment_reply.docx")
        #ExEnd:AddRemoveCommentReply


    def test_process_comments(self):

        #ExStart:ProcessComments
        doc = aw.Document(docs_base.my_dir + "Comments.docx")

        # Extract the information about the comments of all the authors.
        for comment in self.extract_comments(doc):
            print(comment)

        # Remove comments by the "pm" author.
        self.remove_comments_by_author(doc, "pm")
        print('Comments from "pm" are removed!')

        # Extract the information about the comments of the "ks" author.
        for comment in self.extract_comments_by_author(doc, "ks"):
            print(comment)

        # Read the comment's reply and resolve them.
        self.comment_resolved_and_replies(doc)

        # Remove all comments.
        self.remove_comments(doc)
        print("All comments are removed!")

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.process_comments.docx")
        #ExEnd:ProcessComments

    #ExStart:ExtractComments
    @staticmethod
    def extract_comments(doc):

        collected_comments = []
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        for node in comments:
            comment = node.as_comment()
            collected_comments.append(comment.author + " " + comment.date_time.strftime("%Y-%m-%d %H:%M:%S") + " " + comment.to_string(aw.SaveFormat.TEXT))

        return collected_comments

    #ExEnd:ExtractComments

    #ExStart:ExtractCommentsByAuthor
    @staticmethod
    def extract_comments_by_author(doc, author_name):

        collected_comments = []
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        for node in comments:
            comment = node.as_comment()
            if comment.author == author_name:
                collected_comments.append(comment.author + " " + comment.date_time.strftime("%Y-%m-%d %H:%M:%S") + " " + comment.to_string(aw.SaveFormat.TEXT))

        return collected_comments

    #ExEnd:ExtractCommentsByAuthor

    #ExStart:RemoveComments
    @staticmethod
    def remove_comments(doc):

        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        comments.clear()

    #ExEnd:RemoveComments

    #ExStart:RemoveCommentsByAuthor
    @staticmethod
    def remove_comments_by_author(doc, author_name):

        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        # Look through all comments and remove those written by the authorName.
        for i in range(comments.count, 0):
            print(i)
            comment = comments[i].as_comment()
            if comment.author == author_name:
                comment.remove()


    #ExEnd:RemoveCommentsByAuthor


    #ExStart:CommentResolvedandReplies
    @staticmethod
    def comment_resolved_and_replies(doc):

        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        parent_comment = comments[0].as_comment()
        for child in parent_comment.replies:

            child_comment = child.as_comment()
            # Get comment parent and status.
            print(child_comment.ancestor.id)
            print(child_comment.done)

            # And update comment Done mark.
            child_comment.done = True


    #ExEnd:CommentResolvedandReplies

    def test_remove_region_text(self):

        #ExStart:RemoveRegionText
        # Open the document.
        doc = aw.Document(docs_base.my_dir + "Comments.docx")

        comment_start = doc.get_child(aw.NodeType.COMMENT_RANGE_START, 0, True).as_comment_range_start()
        comment_end = doc.get_child(aw.NodeType.COMMENT_RANGE_END, 0, True).as_comment_range_end()

        current_node = comment_start
        is_removing = True
        while current_node is not None and is_removing:
            if current_node.node_type == aw.NodeType.COMMENT_RANGE_END:
                is_removing = False

            next_node = current_node.next_pre_order(doc)
            current_node.remove()
            current_node = next_node

        # Save the document.
        doc.save(docs_base.artifacts_dir + "WorkingWithComments.remove_region_text.docx")
        #ExEnd:RemoveRegionText


if __name__ == '__main__':
    unittest.main()
