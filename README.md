# Food-Log
Food Log allows you to record various food entries by weight in grams, and calculates your DRI / DV / etc. values for the day. 
This is an old VB.NET project I developed for personal use prior to switching to C# and training for certification, hence the lack of input validation, the need for refactoring, etc.
This was only intended to be for private use. I'm only uploading this as an example of a complex program I have written.

Features:
--Keep track of the foods you eat by manually entering the weight, by estimating the weight, or by using the serving sizes that the USDA provides for various foods (such as 1 large apple = 223g) by selecting the size from a combobox
--Generate reports based on these values, such as viewing your DRI values for the day (RDA, or AI if RDA is not available, or no value if neither is available)
--Compare foods or groups of foods to see which offers superior macronurient/vitmain/mineral profiles
--Include or exclude nutrients from reports by going to "Nutrient Editor" tab and checking "Basic" to include or uncheck it to exclude it
--Add new foods to the list
----You can add new foods to the food list in the "Food Editor" tab by clicking the "Add" button on the bottom left
----Enter a name, then enter the USDA site profile (make sure "Full Report" link is used instead of the default "Basic" report)

There are many other features and report types for you to choose from, I just listed some examples.
Screenshots are available in the "Screenshots" folder if you wish to see the contents at a glance.
