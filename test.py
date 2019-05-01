import unittest

import value_tool


class TestResourcePath(unittest.TestCase):
    def test_resource_path(self):
        """
        Tests resource path
        """
        pass


class TestSplash(unittest.TestCase):
    def test_splash_screen(self):
        """
        Tests splash screen
        The layout variable used to define the window is a list with length ==
        amout of elements in the list.

        For the splash screen this length should be 1 as there is only one image 
        in window.
        """
        layout = value_tool.Splash()
        self.assertEqual(len(layout), 1)



if __name__ == "__main__":
    unittest.main()
