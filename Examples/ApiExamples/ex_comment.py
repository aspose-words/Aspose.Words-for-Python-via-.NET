# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from datetime import datetime

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExComment(ApiExampleBase):

    def test_add_comment_with_reply(self):

        #ExStart
        #ExFor:Comment
        #ExFor:Comment.set_text(str)
        #ExFor:Comment.add_reply(str,str,datetime,str)
        #ExSummary:Shows how to add a comment to a document, and then reply to it.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        comment = aw.Comment(doc, "John Doe", "J.D.", datetime.now())
        comment.set_text("My comment.")

        # Place the comment at a node in the document's body.
        # This comment will show up at the location of its paragraph,
        # outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
        builder.current_paragraph.append_child(comment)

        # Add a reply, which will show up under its parent comment.
        comment.add_reply("Joe Bloggs", "J.B.", datetime.now(), "New reply")

        # Comments and replies are both Comment nodes.
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.COMMENT, True).count)

        # Comments that do not reply to other comments are "top-level". They have no ancestor comments.
        self.assertIsNone(comment.ancestor)

        # Replies have an ancestor top-level comment.
        self.assertEqual(comment, comment.replies[0].ancestor)

        doc.save(ARTIFACTS_DIR + "Comment.add_comment_with_reply.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Comment.add_comment_with_reply.docx")
        doc_comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()

        self.assertEqual(1, doc_comment.count)
        self.assertEqual(1, comment.replies.count)

        self.assertEqual("\u0005My comment.\r", doc_comment.get_text())
        self.assertEqual("\u0005New reply\r", doc_comment.replies[0].get_text())

    def test_print_all_comments(self):

        #ExStart
        #ExFor:Comment.ancestor
        #ExFor:Comment.author
        #ExFor:Comment.replies
        #ExFor:CompositeNode.get_child_nodes(NodeType,bool)
        #ExSummary:Shows how to print all of a document's comments and their replies.
        doc = aw.Document(MY_DIR + "Comments.docx")

        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)
        self.assertEqual(12, comments.count) #ExSkip

        # If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
        # Print all top-level comments along with any replies they may have.
        for comment in comments:
            comment = comment.as_comment()
            if comment.ancestor is None:
                print("Top-level comment:")
                print(f"\t\"{comment.get_text().strip()}\", by {comment.author}")
                print(f"Has {comment.replies.count} replies")
                for comment_reply in comment.replies:
                    comment_reply = comment_reply.as_comment()
                    print(f"\t\"{comment_reply.get_text().strip()}\", by {comment_reply.author}")
                print()

        #ExEnd

    def test_remove_comment_replies(self):

        #ExStart
        #ExFor:Comment.remove_all_replies
        #ExFor:Comment.remove_reply(Comment)
        #ExFor:CommentCollection.__getitem__(int)
        #ExSummary:Shows how to remove comment replies.
        doc = aw.Document()

        comment = aw.Comment(doc, "John Doe", "J.D.", datetime.now())
        comment.set_text("My comment.")

        doc.first_section.body.first_paragraph.append_child(comment)

        comment.add_reply("Joe Bloggs", "J.B.", datetime.now(), "New reply")
        comment.add_reply("Joe Bloggs", "J.B.", datetime.now(), "Another reply")

        self.assertEqual(2, comment.replies.count)

        # Below are two ways of removing replies from a comment.
        # 1 -  Use the "remove_reply" method to remove replies from a comment individually:
        comment.remove_reply(comment.replies[0])

        self.assertEqual(1, comment.replies.count)

        # 2 -  Use the "remove_all_replies" method to remove all replies from a comment at once:
        comment.remove_all_replies()

        self.assertEqual(0, comment.replies.count)
        #ExEnd

    def test_done(self):

        #ExStart
        #ExFor:Comment.done
        #ExFor:CommentCollection
        #ExSummary:Shows how to mark a comment as "done".
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Helo world!")

        # Insert a comment to point out an error.
        comment = aw.Comment(doc, "John Doe", "J.D.", datetime.now())
        comment.set_text("Fix the spelling error!")
        doc.first_section.body.first_paragraph.append_child(comment)

        # Comments have a "done" flag, which is set to "False" by default.
        # If a comment suggests that we make a change within the document,
        # we can apply the change, and then also set the "done" flag afterwards to indicate the correction.
        self.assertFalse(comment.done)

        doc.first_section.body.first_paragraph.runs[0].text = "Hello world!"
        comment.done = True

        # Comments that are "done" will differentiate themselves
        # from ones that are not "done" with a faded text color.
        comment = aw.Comment(doc, "John Doe", "J.D.", datetime.now())
        comment.set_text("Add text to this paragraph.")
        builder.current_paragraph.append_child(comment)

        doc.save(ARTIFACTS_DIR + "Comment.done.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Comment.done.docx")
        comment = doc.get_child_nodes(aw.NodeType.COMMENT, True)[0].as_comment()

        self.assertTrue(comment.done)
        self.assertEqual("\u0005Fix the spelling error!", comment.get_text().strip())
        self.assertEqual("Hello world!", doc.first_section.body.first_paragraph.runs[0].text)

    ##ExStart
    ##ExFor:Comment.done
    ##ExFor:Comment.__init__(DocumentBase)
    ##ExFor:Comment.accept(DocumentVisitor)
    ##ExFor:Comment.date_time
    ##ExFor:Comment.id
    ##ExFor:Comment.initial
    ##ExFor:CommentRangeEnd
    ##ExFor:CommentRangeEnd.__init__(DocumentBase,int)
    ##ExFor:CommentRangeEnd.accept(DocumentVisitor)
    ##ExFor:CommentRangeEnd.id
    ##ExFor:CommentRangeStart
    ##ExFor:CommentRangeStart.__init__(DocumentBase,int)
    ##ExFor:CommentRangeStart.accept(DocumentVisitor)
    ##ExFor:CommentRangeStart.id
    ##ExSummary:Shows how print the contents of all comments and their comment ranges using a document visitor.
    #def test_create_comments_and_print_all_info(self):

    #    doc = aw.Document()
    #    new_comment = aw.Comment(doc)
    #    new_comment.author = "VDeryushev"
    #    new_comment.initial = "VD"
    #    new_comment.date_time = datetime.now()

    #    new_comment.set_text("Comment regarding text.")

    #    # Add text to the document, warp it in a comment range, and then add your comment.
    #    para = doc.first_section.body.first_paragraph
    #    para.append_child(aw.CommentRangeStart(doc, new_comment.id))
    #    para.append_child(aw.Run(doc, "Commented text."))
    #    para.append_child(aw.CommentRangeEnd(doc, new_comment.id))
    #    para.append_child(new_comment)

    #    # Add two replies to the comment.
    #    new_comment.add_reply("John Doe", "JD", datetime.now(), "New reply.")
    #    new_comment.add_reply("John Doe", "JD", datetime.now(), "Another reply.")

    #    ExComment.print_all_comment_info(doc.get_child_nodes(aw.NodeType.COMMENT, True))

    #def print_all_comment_info(comments: aw.NodeCollection):
    #    """Iterates over every top-level comment and prints its comment range, contents, and replies."""

    #    comment_visitor = aw.CommentInfoPrinter()

    #    # Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
    #    for comment in comments:
    #        comment = comment.as_comment()
    #        if comment.ancestor is None:
    #            # First, visit the start of the comment range.
    #            comment_range_start = comment.previous_sibling.previous_sibling.previous_sibling.as_comment_range_start()
    #            comment_range_start.accept(comment_visitor)

    #            # Then, visit the comment, and any replies that it may have.
    #            comment.accept(comment_visitor)

    #            for reply in comment.replies:
    #                reply.accept(comment_visitor)

    #            # Finally, visit the end of the comment range, and then print the visitor's text contents.
    #            comment_range_end = comment.previous_sibling.as_comment_range_end()
    #            comment_range_end.accept(comment_visitor)

    #            print(comment_visitor.get_text())

    #class CommentInfoPrinter(aw.DocumentVisitor):
    #    """Prints information and contents of all comments and comment ranges encountered in the document."""

    #    def __init__(self):

    #        self.builder = io.StringIO()
    #        self.visitor_is_inside_comment = False
    #        self.doc_traversal_depth = 0

    #    def get_text() -> str:
    #        """Gets the plain text of the document that was accumulated by the visitor."""

    #        return self.builder.getvalue()

    #    def visit_run(self, run: aw.Run) -> aw.VisitorAction:
    #        """Called when a Run node is encountered in the document."""

    #        if self.visitor_is_inside_comment:
    #           self.indent_and_append_line("[Run] \"" + run.text + "\"")

    #        return aw.VisitorAction.CONTINUE

    #    def visit_comment_range_start(self, comment_range_start: aw.CommentRangeStart) -> aw.VisitorAction:
    #        """Called when a CommentRangeStart node is encountered in the document."""

    #        self.indent_and_append_line("[Comment range start] ID: " + comment_range_start.id)
    #        self.doc_traversal_depth += 1
    #        self.visitor_is_inside_comment = True

    #        return aw.VisitorAction.CONTINUE

    #    def visit_comment_range_end(self, comment_range_end: aw.CommentRangeEnd) -> aw.VisitorAction:
    #        """Called when a CommentRangeEnd node is encountered in the document."""

    #        self.doc_traversal_depth -= 1
    #        self.indent_and_append_line("[Comment range end] ID: " + comment_range_end.id + "\n")
    #        self.visitor_is_inside_comment = False

    #        return aw.VisitorAction.CONTINUE

    #    def visit_comment_start(self, comment: aw.Comment) -> aw.VisitorAction:
    #        """Called when a Comment node is encountered in the document."""

    #        self.indent_and_append_line(
    #            f"[Comment start] For comment range ID {comment.id}, By {comment.author} on {comment.date_time}")
    #        self.doc_traversal_depth += 1
    #        self.visitor_is_inside_comment = True

    #        return aw.VisitorAction.CONTINUE

    #    def visit_comment_end(self, comment: aw.Comment) -> aw.VisitorAction:
    #        """Called when the visiting of a Comment node is ended in the document."""

    #        self.doc_traversal_depth -= 1
    #        self.indent_and_append_line("[Comment end]")
    #        self.visitor_is_inside_comment = False

    #        return aw.VisitorAction.CONTINUE

    #    def indent_and_append_line(self, text: str):
    #        """Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree."""

    #        self.builder.write("|  " * self.doc_traversal_depth)
    #        self.builder.write(text + "\n")

    ##ExEnd
