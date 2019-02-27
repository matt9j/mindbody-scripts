# mindbody-scripts
A collection of scripts for automating tasks associated with running a
yoga studio with Mindbody

## splitter

Splitter splits a payroll report up by instructor for inclusion in the
instructor's payroll email. If you have an exported excel payroll
report called full_report.xlsx, you can use splitter like this:
`python splitter --pay_period "Feb 1989" full_report.xlsx`. This will
generate individual reports titled "[Instructor Name] - Feb 1989.xlsx"
The name format can be changed by modifying the splitter program.
