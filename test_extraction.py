import unittest
from app import LinkExtractor

class TestLinkExtraction(unittest.TestCase):
    
    def setUp(self):
        self.extractor = LinkExtractor()
        self.test_url = "https://help.salesforce.com/s/articleView?id=release-notes.rn_permissions.htm&release=254&type=5"
    
    def test_permissions_page_extraction(self):
        """Test that we extract only relevant release note content links, not all page links"""
        
        # Expected links from the permissions release notes page
        expected_links = [
            "Delivered Idea: Manage Included Permission Sets in Permission Set Groups via Summaries",
            "Delivered Idea: Allow Users to View All Fields for a Specified Object", 
            "The View All and Modify All Object Permissions Have New Names",
            "Remove User and Custom Permissions in Permission Set Summaries"
        ]
        
        # Extract links
        links, page_title = self.extractor.extract_links_from_url(self.test_url)
        
        # Assertions
        self.assertEqual(page_title, "Permissions")
        
        # Should extract only relevant content links, not 1500+ navigation links
        self.assertLess(len(links), 20, f"Too many links extracted: {len(links)}. Expected < 20 content links.")
        self.assertGreater(len(links), 0, "No links were extracted")
        
        # Check that expected links are found
        extracted_texts = [link['text'] for link in links]
        
        for expected_link in expected_links:
            self.assertIn(expected_link, extracted_texts, 
                         f"Expected link '{expected_link}' not found in extracted links: {extracted_texts}")
        
        # Ensure we're not getting navigation/footer links
        unwanted_patterns = ['Home', 'Login', 'Support', 'Contact', 'Privacy', 'Terms', 'Footer', 'Navigation']
        for link in links:
            for pattern in unwanted_patterns:
                self.assertNotIn(pattern.lower(), link['text'].lower(), 
                               f"Unwanted navigation link found: {link['text']}")
        
        # All links should be release note related
        for link in links:
            self.assertTrue(
                'release-notes' in link['url'] or 'articleView' in link['url'],
                f"Non-release note link found: {link['url']}"
            )
    
    def test_list_views_page_extraction(self):
        """Test extraction from list views release notes page"""
        
        list_views_url = "https://help.salesforce.com/s/articleView?id=release-notes.rn_list_views.htm&release=254&type=5"
        
        links, page_title = self.extractor.extract_links_from_url(list_views_url)
        
        # Should extract only relevant content links
        self.assertLess(len(links), 20, f"Too many links extracted: {len(links)}")
        self.assertGreater(len(links), 0, "No links were extracted")
        
        # Title should be meaningful, not generic
        self.assertNotEqual(page_title, "Help And Training Community")
        self.assertIn("List Views", page_title)

if __name__ == '__main__':
    unittest.main() 