# The SQL Job Blender 🌀

### Combine your SQL Agent jobs into a single smoothie of efficiency! 🍹

Ever felt like your SQL Agent jobs were scattered like puzzle pieces? Introducing the SQL Job Blender – your one-stop solution to merge multiple jobs into a single, streamlined process. Say goodbye to job chaos and hello to job harmony!

### What Does it Do?

This script takes those scattered SQL Agent jobs, throws them into a metaphorical blender, and voilà – out comes a delicious concoction of a single job. It's like making a smoothie, but for your SQL Server!

### How to Use:

1. **Ingredients:**
   - Set `@NewJobName` to your desired name for the new, combined job.
   - Specify `@SourceJobNameSpec` to filter out the jobs you want to combine.
   - Give a catchy name to your job's schedule with `@ScheduleName`.

2. **Blend It:**
   - Run the script and watch as it combines the steps from your specified jobs into one delicious blend of efficiency.

3. **Serve Daily:**
   - With a schedule set at 7:15 AM daily, your job smoothie will be ready to serve like clockwork.

### But Wait, There's More!

- **Cleanup Included:**
  We're not leaving a mess behind. The script also cleans up after itself, dropping the temporary table like a responsible bartender cleaning up the bar.

- **Customize to Taste:**
  Feel free to tweak the script to fit your unique flavor of SQL job management. Add your favorite garnishes, adjust the schedule – make it your own!

So why settle for a job salad when you can have a job smoothie? Try the SQL Job Blender today and blend your way to SQL success! 🎉
