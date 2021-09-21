import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithComments(docs_base.DocsExamplesBase):
    
    def test_add_comments(self) :
        
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
        

    def test_anchor_comment(self) :
        
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

        commentRangeStart = aw.CommentRangeStart(doc, comment.id)
        commentRangeEnd = aw.CommentRangeEnd(doc, comment.id)

        run1.parent_node.insert_after(commentRangeStart, run1)
        run3.parent_node.insert_after(commentRangeEnd, run3)
        commentRangeEnd.parent_node.insert_after(comment, commentRangeEnd)

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.anchor_comment.doc")
        #ExEnd:AnchorComment
        

    def test_add_remove_comment_reply(self) :
        
        #ExStart:AddRemoveCommentReply
        doc = aw.Document(docs_base.my_dir + "Comments.docx")

        comment = doc.get_child(aw.NodeType.COMMENT, 0, True).as_comment()
        comment.remove_reply(comment.replies[0])

        comment.add_reply("John Doe", "JD", datetime(2017, 9, 25, 12, 15, 0), "New reply")

        doc.save(docs_base.artifacts_dir + "WorkingWithComments.add_remove_comment_reply.docx")
        #ExEnd:AddRemoveCommentReply
        

    def test_process_comments(self) :
        
        #ExStart:ProcessComments
        doc = aw.Document(docs_base.my_dir + "Comments.docx")

        # Extract the information about the comments of all the authors.
        for comment in self.extract_comments(doc) :
            print(comment)

        # Remove comments by the "pm" author.
        self.remove_comments_by_author(doc, "pm")
        print("Comments from \"pm\" are removed!")

        # Extract the information about the comments of the "ks" author.
        for comment in self.extract_comments_by_author(doc, "ks") :
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
    def extract_comments(doc) :
        
        collectedComments = []
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        for node in comments :
            comment = node.as_comment()
            collectedComments.append(comment.author + " " + comment.date_time.strftime("%Y-%m-%d %H:%M:%S") + " " + comment.to_string(aw.SaveFormat.TEXT))
            
        return collectedComments
        
    #ExEnd:ExtractComments
        
    #ExStart:ExtractCommentsByAuthor
    @staticmethod
    def extract_comments_by_author(doc, authorName) :
        
        collectedComments = []
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        for node in comments :
            comment = node.as_comment()
            if (comment.author == authorName) :
                collectedComments.append(comment.author + " " + comment.date_time.strftime("%Y-%m-%d %H:%M:%S") + " " + comment.to_string(aw.SaveFormat.TEXT))
            
        return collectedComments
        
    #ExEnd:ExtractCommentsByAuthor

    #ExStart:RemoveComments
    @staticmethod
    def remove_comments(doc) :
        
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        comments.clear()
        
    #ExEnd:RemoveComments

    #ExStart:RemoveCommentsByAuthor
    @staticmethod
    def remove_comments_by_author(doc, authorName) :
        
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        # Look through all comments and remove those written by the authorName.
        for i in range(comments.count, 0) :
            print(i)
            comment = comments[i].as_comment()
            if (comment.author == authorName) :
                comment.remove()
            
        
    #ExEnd:RemoveCommentsByAuthor


    #ExStart:CommentResolvedandReplies
    @staticmethod
    def comment_resolved_and_replies(doc) :
        
        comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)

        parentComment = comments[0].as_comment()
        for child in parentComment.replies :
            
            childComment = child.as_comment()
            # Get comment parent and status.
            print(childComment.ancestor.id)
            print(childComment.done)

            # And update comment Done mark.
            childComment.done = True
            
        
    #ExEnd:CommentResolvedandReplies
    


if __name__ == '__main__':
    unittest.main()