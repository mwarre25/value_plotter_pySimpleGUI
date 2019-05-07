import unittest

import value_tool

class TestValueLookup(unittest.TestCase):
    def test_value_lookup(self):
        """Tests value lookup table
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

    def test_not_value_lookup(self):
        """
        Tests value lookup table against incorrect values
        """
        wrong_lookups_dict = {
            (1, 1): 6,
            (1, 2): 6,
            (1, 3): 7,
            (1, 4): 8,
            (1, 5): 8,
            (2, 1): 6,
            (2, 2): 6,
            (2, 3): 7,
            (2, 4): 9,
            (2, 5): 9,
            (3, 1): 6,
            (3, 2): 7,
            (3, 3): 8,
            (3, 4): 10,
            (3, 5): 10,
            (4, 1): 7,
            (4, 2): 8,
            (4, 3): 9,
            (4, 4): 10,
            (4, 5): 10,
            (5, 1): 7,
            (5, 2): 9,
            (5, 3): 10,
            (5, 4): 10,
            (5, 5): 10,
        }

        for key, value in wrong_lookups_dict.items():
            self.assertNotEqual(value_tool.value_lookup(key), value)

if __name__ == "__main__":
    unittest.main()
