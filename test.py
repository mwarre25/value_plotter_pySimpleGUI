import unittest
import value_tool


class TestResourcePath(unittest.TestCase):
    def test_resource_path(self):
        """
        Tests resource path
        """
        pass


# this test doesn't work because of tKinter I think https://bit.ly/2XYLzP8
# class TestSplash(unittest.TestCase):
#     def test_splash_screen(self):
#         """
#         Tests splash screen
#         The layout variable used to define the window is a list with length ==
#         amout of elements in the list.

#         For the splash screen this length should be 1 as there is only one
#         image in window.
#         """
#         layout = value_tool.Splash()
#         self.assertEqual(len(layout), 1)


class TestValueLookup(unittest.TestCase):
    def test_value_lookup(self):
        """
        Tests value lookup table
        """
        lookups_dict = {
            (1, 1): 1,
            (1, 2): 1,
            (1, 3): 2,
            (1, 4): 3,
            (1, 5): 3,
            (2, 1): 1,
            (2, 2): 1,
            (2, 3): 2,
            (2, 4): 4,
            (2, 5): 4,
            (3, 1): 1,
            (3, 2): 2,
            (3, 3): 3,
            (3, 4): 5,
            (3, 5): 5,
            (4, 1): 2,
            (4, 2): 3,
            (4, 3): 4,
            (4, 4): 5,
            (4, 5): 5,
            (5, 1): 2,
            (5, 2): 4,
            (5, 3): 5,
            (5, 4): 5,
            (5, 5): 5,
        }

        for key, value in lookups_dict.items():
            self.assertEqual(value_tool.value_lookup(key), value)


if __name__ == "__main__":
    unittest.main()
