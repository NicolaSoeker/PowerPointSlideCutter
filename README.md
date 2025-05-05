## PowerPointSlideCutter

Cuts a Power Point presentation into seperate slides. 


- The script will create a folder called "Split_Slides" in the same location as your PowerPoint file.
  Inside that folder, you'll find each slide saved as a separate .pptx file.
- The name of each created slide will be the content of the top left pptx shape within a certain margin (less than 100px from top and less than 200px from left)



## Make macro persistent


You save the macro-enabled template as a PowerPoint Add-in (*.ppam). PowerPoint should automatically place this in AppData\Microsoft\AddIns in your user folder. Then you load it into PowerPoint using File>Options>Add-ins. Change the Manage dropdown to PowerPoint Add-ins and click on the Go button. In the Add-ins dialog, click on Add New, then select your addin and click on the Open button. Back in the Add-ins dialog, double-check that your add-in has a check mark beside it, then OK out.

Once you've set it up, the add-in will automatically load every time you start PowerPoint and your macros will be available to all presentations.


