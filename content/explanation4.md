# Explanation of Fields

Below this explanation box, you may find a list of fields that have been filled in. The headers of these fields may be difficult to understand, so this explanation box is provided to clarify the meaning of each header.

## iSegments

The `iSegments` field lists the segments after they have been mapped or reordered. These segments are correctly positioned and will be included in the output file of Juyo in this order. It is important to carefully check that the segments are in the correct order.

## iTerm

The `iTerm` field contains the terminology for room nights and revenue or ADR. It is important to ensure that this information is entered correctly, as the text and the cell value must be EXACTLY the same.

## iSort

The `iSort` field lists the segments as they were entered. The script uses the `iSort` field to sort the segments, so it is important to make sure that these segments are in the exact order as they appear in your Excel file.

## iSkipper_s and iSkipper_s1

The `iSkipper_s` and `iSkipper_s1` fields are used if you want to skip certain terminology (see the explanation box on "Skip Terminology"). These fields can be left empty if desired.

## iLoc

The `iLoc` field is used to indicate the row or column where the terminology is stored. A value of `0` indicates that the terminology is stored in a row or column, the value of `2` indicates in which row or column it is stored.

# Important to read:
It is important to carefully check that all of the fields are correct and to make any necessary changes. After making any necessary changes, run the conversion process and check the output before saving the data for future use. This will ensure that the saved data is correct.
