"""Chapter 3 practice: Python and SciPy crash course.

This file follows Chapter 3 of *Machine Learning Mastery With Python* and
contains small examples of Python, NumPy, Matplotlib, and Pandas.

The examples are intentionally kept independent from the Titanic project.
They are written for Python 3 and are ready to run when the dependencies are
available.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def python_crash_course() -> None:
    """Practice assignment, flow control, data structures, and functions."""
    # Assignment and basic types.
    text = "hello world"
    value = 123.1
    is_ready = True
    empty_value = None
    first, second, third = 1, 2, 3

    print(text[0])
    print(len(text))
    print(value, is_ready, empty_value)
    print(first, second, third)

    # If/elif/else.
    speed = 99
    if speed == 99:
        print("That is fast")
    elif speed > 200:
        print("That is too fast")
    else:
        print("That is safe")

    # For-loop and while-loop.
    for number in range(3):
        print("for:", number)

    number = 0
    while number < 3:
        print("while:", number)
        number += 1

    # Tuple, list, and dictionary.
    coordinates = (1, 2, 3)
    values = [1, 2, 3]
    values.append(4)
    metadata = {"a": 1, "b": 2, "c": 3}
    metadata["a"] = 11

    print("tuple:", coordinates)
    print("list:", values)
    print("dictionary keys:", list(metadata.keys()))
    print("dictionary values:", list(metadata.values()))

    def add(left: float, right: float) -> float:
        return left + right

    print("function result:", add(1, 3))


def numpy_crash_course() -> None:
    """Practice array creation, indexing, shapes, and arithmetic."""
    one_dimensional = np.array([1, 2, 3])
    two_dimensional = np.array([[1, 2, 3], [3, 4, 5]])

    print("array:", one_dimensional)
    print("shape:", one_dimensional.shape)
    print("matrix:\n", two_dimensional)
    print("matrix shape:", two_dimensional.shape)
    print("first row:", two_dimensional[0])
    print("last row:", two_dimensional[-1])
    print("specific value:", two_dimensional[0, 2])
    print("whole column:", two_dimensional[:, 2])

    first_array = np.array([2, 2, 2])
    second_array = np.array([3, 3, 3])
    print("addition:", first_array + second_array)
    print("multiplication:", first_array * second_array)


def matplotlib_crash_course() -> None:
    """Create the line and scatter plots introduced in Chapter 3."""
    values = np.array([1, 2, 3])

    plt.figure(figsize=(8, 4))
    plt.subplot(1, 2, 1)
    plt.plot(values)
    plt.xlabel("x axis")
    plt.ylabel("y axis")
    plt.title("Line plot")

    x_values = np.array([1, 2, 3])
    y_values = np.array([2, 4, 6])
    plt.subplot(1, 2, 2)
    plt.scatter(x_values, y_values)
    plt.xlabel("x axis")
    plt.ylabel("y axis")
    plt.title("Scatter plot")

    plt.tight_layout()
    plt.show()


def pandas_crash_course() -> None:
    """Practice Series and DataFrame creation and column access."""
    series = pd.Series(np.array([1, 2, 3]), index=["a", "b", "c"])
    print("series:\n", series)
    print("series by position:", series.iloc[0])
    print("series by label:", series.loc["a"])

    frame = pd.DataFrame(
        np.array([[1, 2, 3], [4, 5, 6]]),
        index=["a", "b"],
        columns=["one", "two", "three"],
    )
    print("dataframe:\n", frame)
    print("column by name:\n", frame["one"])
    print("column by attribute:\n", frame.one)


def main() -> None:
    """Run all Chapter 3 examples in order."""
    python_crash_course()
    numpy_crash_course()
    pandas_crash_course()
    matplotlib_crash_course()


if __name__ == "__main__":
    main()
