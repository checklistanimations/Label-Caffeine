# Label Caffeine
a VBA framework that can be added into projects to provide modern day button functionality using labels and the core code.


# Features
* Easy Mouse Hover Events: Removes the need to write out separate mouse up, mouse down events.
* Glass Triggers: Provides mouse hover events to unsupported controls like spins and scrolls.
* Linked Controls: Hover events for multiple controls simultaneously for something like a custom checkbox with 2 or more controls.
* Visual Styling: Create the style on a separate control and reference it. No need to type out RGB colors or other properties.
* Better Control Feedback: Push away buttons, font color changes, Even old school special effect updates all easily customizable for click events.
* Ability to Toggle Tooltips On and Off: Provides a way to disable ControlTipText that would overlap UI with a simple toggle button.
* Minimal User Code Required: All contained in a class and a module, users can make minimal code updates in order to give their userform a modern day feel.


# Concept Workflow (Under Construction):
The line of setupLabelCaffeine(frm) will be added to a form initialize event. The form pointer will be passed in. 
This setup sub is located in the module of label caffeine. There are 3 module level variables in this module. 
* LC_PREFIX (static string defaults to LC==) 
* setup Fail Level (integer) 
* LastHovered (string) 

When setup runs it opens a pipeline that creates a new collection. The collection will be referred as recs. Recs is then passed to the first part of the pipeline. 

### HARVEST CONTROLS 
This will loop through all controls on the form and check the tag. If the left of the tag matches the prefix, the control will be added to the collection. 
The collection holds the control name for the key which will prevent duplicates 
The value of the collection is an array with the following indexes. 
* control name (can't pull key jn collections)
* Control type
* control tag in it's exact state
* Clean tag - done later 
* Errors - blank string for now
If a control does now have a qualifying tag with the prefix it will not be used in LC. This includes validation checks and other functionality. 
At the end of this sub we should have only the collection created with the string array. 
If the collection is empty set error level to  9 and exit pipeline. Setup will check error level and present message box of no valid controls found. LC will end and recs will be cleared. (set recs = nothing). This will probably be a special exit function since we need to destroy the form object as well. . 

### Validate and Sanitize Tag
The pipeline will look at each tag and verify.... 







 A static variable of LC_PREFIX will indicate the prefix to check and can be changed. By default it is LC==

* Verify prefix
* Check for tag formatting
* check for duplicate keys of useBack = 1 and useBack = 0 in same tag
* check for incompatible properties per type
* store control name, type, tag, clean tag and errors in array of collection
* check for unknown keys
* Check that all groups have at least 2 members
* check that stylefroms exist on form / in collection
* create a simple key exists in collection function
* clean tag by fixing prefix to ; then split all by ;



### Properties that can be used in LC tags
- useBack
- useOpacity
- useFore
- useFont
- useOutline
- useBorder
- useSpecial
- usePicture
- useGlass
- group
- action
- tooltip



### FAQ



### Checklist
- [ ] finish documentation
- [ ] figure out new pipeline
- [ ] see if checklist works
- [ ] do we need the space



Explanation
LabelCaffeine: = The identifier
action = The sub that will run when the control is clicked
useBack = true or false with 1 or 0. If 1 then background color comes from styleFrom on hover
useOutline = same as useBack but border color
useFore = same as useBack but fore color (font color)
useFont = same as useBack but full font style with bold, family italic etc.
useImage = same as useBack but image pointer of the styleFrom control image
useSpecial = same as useBack but the specialEffect
styleFrom = The control on the userform that has the properties to use when hovering a control. User can have as many styleFrom controls as they want.
Tooltip = A way to store controltiptext so the user does not always have a tooltip appearing. It can be enabled and disabled with a button on their form that will call labCafTooltipsOn or labCafTooltipsOff
