# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from datetime import datetime, timezone
import aspose.words as aw
import unittest
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
        comment = aw.Comment(doc, 'John Doe', 'J.D.', datetime.now())
        comment.set_text('My comment.')
        # Place the comment at a node in the document's body.
        # This comment will show up at the location of its paragraph,
        # outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
        builder.current_paragraph.append_child(comment)
        # Add a reply, which will show up under its parent comment.
        comment.add_reply('Joe Bloggs', 'J.B.', datetime.now(), 'New reply')
        # Comments and replies are both Comment nodes.
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.COMMENT, True).count)
        # Comments that do not reply to other comments are "top-level". They have no ancestor comments.
        self.assertIsNone(comment.ancestor)
        # Replies have an ancestor top-level comment.
        self.assertEqual(comment, comment.replies[0].ancestor)
        doc.save(ARTIFACTS_DIR + 'Comment.add_comment_with_reply.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Comment.add_comment_with_reply.docx')
        doc_comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()
        self.assertEqual(1, doc_comment.count)
        self.assertEqual(1, comment.replies.count)
        self.assertEqual('\x05My comment.\r', doc_comment.get_text())
        self.assertEqual('\x05New reply\r', doc_comment.replies[0].get_text())

    def test_print_all_comments(self):
        #ExStart
        #ExFor:Comment.ancestor
        #ExFor:Comment.author
        #ExFor:Comment.replies
        #ExFor:CompositeNode.get_child_nodes(NodeType,bool)
        #ExSummary:Shows how to print all of a document's comments and their replies.
        doc = aw.Document(MY_DIR + 'Comments.docx')
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)
        self.assertEqual(12, comments.count)  #ExSkip
        # If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
        # Print all top-level comments along with any replies they may have.
        for comment in comments:
            comment = comment.as_comment()
            if comment.ancestor is None:
                print('Top-level comment:')
                print(f'\t"{comment.get_text().strip()}", by {comment.author}')
                print(f'Has {comment.replies.count} replies')
                for comment_reply in comment.replies:
                    comment_reply = comment_reply.as_comment()
                    print(f'\t"{comment_reply.get_text().strip()}", by {comment_reply.author}')
                print()
        #ExEnd

    def test_remove_comment_replies(self):
        #ExStart
        #ExFor:Comment.remove_all_replies
        #ExFor:Comment.remove_reply(Comment)
        #ExFor:CommentCollection.__getitem__(int)
        #ExSummary:Shows how to remove comment replies.
        doc = aw.Document()
        comment = aw.Comment(doc, 'John Doe', 'J.D.', datetime.now())
        comment.set_text('My comment.')
        doc.first_section.body.first_paragraph.append_child(comment)
        comment.add_reply('Joe Bloggs', 'J.B.', datetime.now(), 'New reply')
        comment.add_reply('Joe Bloggs', 'J.B.', datetime.now(), 'Another reply')
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
        builder.writeln('Helo world!')
        # Insert a comment to point out an error.
        comment = aw.Comment(doc, 'John Doe', 'J.D.', datetime.now())
        comment.set_text('Fix the spelling error!')
        doc.first_section.body.first_paragraph.append_child(comment)
        # Comments have a "done" flag, which is set to "False" by default.
        # If a comment suggests that we make a change within the document,
        # we can apply the change, and then also set the "done" flag afterwards to indicate the correction.
        self.assertFalse(comment.done)
        doc.first_section.body.first_paragraph.runs[0].text = 'Hello world!'
        comment.done = True
        # Comments that are "done" will differentiate themselves
        # from ones that are not "done" with a faded text color.
        comment = aw.Comment(doc, 'John Doe', 'J.D.', datetime.now())
        comment.set_text('Add text to this paragraph.')
        builder.current_paragraph.append_child(comment)
        doc.save(ARTIFACTS_DIR + 'Comment.done.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Comment.done.docx')
        comment = doc.get_child_nodes(aw.NodeType.COMMENT, True)[0].as_comment()
        self.assertTrue(comment.done)
        self.assertEqual('\x05Fix the spelling error!', comment.get_text().strip())
        self.assertEqual('Hello world!', doc.first_section.body.first_paragraph.runs[0].text)

    def test_utc_date_time(self):
        #ExStart:UtcDateTime
        #ExFor:Comment.date_time_utc
        #ExSummary:Shows how to get UTC date and time.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        date = datetime.now()
        comment = aw.Comment(doc, 'John Doe', 'J.D.', date)
        comment.set_text('My comment.')
        builder.current_paragraph.append_child(comment)
        doc.save(file_name=ARTIFACTS_DIR + 'Comment.UtcDateTime.docx')

        doc = aw.Document(ARTIFACTS_DIR + 'Comment.UtcDateTime.docx')
        comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()

        # DateTimeUtc return data without milliseconds.
        self.assertEqual(date.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"), comment.date_time_utc.strftime("%Y-%m-%d %H:%M:%S"))
        #ExEnd:UtcDateTime
