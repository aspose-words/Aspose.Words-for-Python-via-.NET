import unittest
from datetime import date, datetime

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExComment(aeb.ApiExampleBase):
    
    def test_add_comment_with_reply(self) :
        
        #ExStart
        #ExFor:Comment
        #ExFor:Comment.set_text(String)
        #ExFor:Comment.add_reply(String, String, DateTime, String)
        #ExSummary:Shows how to add a comment to a document, and then reply to it.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        comment = aw.Comment(doc, "John Doe", "J.d.", date.today())
        comment.set_text("My comment.")
            
        # Place the comment at a node in the document's body.
        # This comment will show up at the location of its paragraph,
        # outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
        builder.current_paragraph.append_child(comment)

        # Add a reply, which will show up under its parent comment.
        comment.add_reply("Joe Bloggs", "J.b.", date.today(), "New reply")

        # Comments and replies are both Comment nodes.
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.COMMENT, True).count)

        # Comments that do not reply to other comments are "top-level". They have no ancestor comments.
        self.assertIsNone(comment.ancestor)

        # Replies have an ancestor top-level comment.
        self.assertEqual(comment, comment.replies[0].ancestor)

        doc.save(aeb.artifacts_dir + "Comment.add_comment_with_reply.docx")
        #ExEnd

# there is no casting yet.
#        doc = new Document(aeb.artifacts_dir + "Comment.add_comment_with_reply.docx")
#        Comment docComment = (Comment)doc.get_child(NodeType.comment, 0, true)
#
#        self.assertEqual(1, docComment.count)
#        self.assertEqual(1, comment.replies.count)
#
#        self.assertEqual("\u0005My comment.\r", docComment.get_text())
#        self.assertEqual("\u0005New reply\r", docComment.replies[0].get_text())
        

    def test_print_all_comments(self) :
        
        #ExStart
        #ExFor:Comment.ancestor
        #ExFor:Comment.author
        #ExFor:Comment.replies
        #ExFor:CompositeNode.get_child_nodes(NodeType, Boolean)
        #ExSummary:Shows how to print all of a document's comments and their replies.
        doc = aw.Document(aeb.my_dir + "Comments.docx")

        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)
        self.assertEqual(12, comments.count) #ExSkip

        # If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
        # Print all top-level comments along with any replies they may have.
        for comment in comments:
            comment = comment.as_comment()
            if comment.ancestor == None :
                print("Top-level comment:")
                print(f"\t\"{comment.get_text().strip()}\", by comment.author")
                print(f"Has {comment.replies.count} replies")
                for comment_reply in comment.replies:
                    comment_reply = comment_reply.as_comment()
                    print(f"\t\"{comment_reply.get_text().strip()}\", by {comment_reply.author}")
                    
                print()
            
        #ExEnd
        

    def test_remove_comment_replies(self) :
        
        #ExStart
        #ExFor:Comment.remove_all_replies
        #ExFor:Comment.remove_reply(Comment)
        #ExFor:CommentCollection.item(Int32)
        #ExSummary:Shows how to remove comment replies.
        doc = aw.Document()

        comment = aw.Comment(doc, "John Doe", "J.d.", date.today())
        comment.set_text("My comment.")

        doc.first_section.body.first_paragraph.append_child(comment)
            
        comment.add_reply("Joe Bloggs", "J.b.", date.today(), "New reply")
        comment.add_reply("Joe Bloggs", "J.b.", date.today(), "Another reply")

        self.assertEqual(2, comment.replies.count) 

        # Below are two ways of removing replies from a comment.
        # 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
        comment.remove_reply(comment.replies[0])

        self.assertEqual(1, comment.replies.count)

        # 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
        comment.remove_all_replies()

        self.assertEqual(0, comment.replies.count) 
        #ExEnd
        

    def test_done(self) :
        
        #ExStart
        #ExFor:Comment.done
        #ExFor:CommentCollection
        #ExSummary:Shows how to mark a comment as "done".
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Helo world!")

        # Insert a comment to point out an error. 
        comment = aw.Comment(doc, "John Doe", "J.d.", date.today())
        comment.set_text("Fix the spelling error!")
        doc.first_section.body.first_paragraph.append_child(comment)

        # Comments have a "Done" flag, which is set to "false" by default. 
        # If a comment suggests that we make a change within the document,
        # we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
        self.assertFalse(comment.done)

        doc.first_section.body.first_paragraph.runs[0].text = "Hello world!"
        comment.done = True

        # Comments that are "done" will differentiate themselves
        # from ones that are not "done" with a faded text color.
        comment = aw.Comment(doc, "John Doe", "J.d.", date.today())
        comment.set_text("Add text to this paragraph.")
        builder.current_paragraph.append_child(comment)

        doc.save(aeb.artifacts_dir + "Comment.done.docx")
        #ExEnd

# there is no type casting yet.
#        doc = new Document(aeb.artifacts_dir + "Comment.done.docx")
#        comment = (Comment)doc.get_child_nodes(NodeType.comment, true)[0]
#
#        self.assertTrue(comment.done)
#        self.assertEqual("\u0005Fix the spelling error!", comment.get_text().strip())
#        self.assertEqual("Hello world!", doc.first_section.body.first_paragraph.runs[0].text)
        
        
   
if __name__ == '__main__':
    unittest.main()    