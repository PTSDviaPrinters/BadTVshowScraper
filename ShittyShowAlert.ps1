
# Start the web Scraper useing Firefox
if (!(Get-InstalledModule Selenium)) {
    Install-Module Selenium
}
else {
    Write-Host "Module Selenim is present"
}
if (!(Get-InstalledModule ImportExcel)) {
    Install-Module -Name ImportExcel
}
else {
    Write-Host "Module ImportExcel is present"
}
$Driver = Start-SeFirefox -Headless
Enter-SeUrl 'https://www.tvguide.com/listings/' -Driver $Driver

    # The comand below gives the output below that.
    # $Driver.FindElementByXPath('//*[@id="listings"]/ul/li[X]')

    # WrappedDriver                        : OpenQA.Selenium.Firefox.FirefoxDriver
    # TagName                              : span
    # Text                                 : The Shining12:00 PM - 3:30 PM
    # Enabled                              : True
    # Selected                             : False
    # Location                             : {X=225,Y=1451}
    # Size                                 : {Width=1041, Height=34}
    # Displayed                            : True
    # LocationOnScreenOnceScrolledIntoView : {X=225,Y=941}
    # Coordinates                          : OpenQA.Selenium.Remote.RemoteCoordinates

# Here I am makeing a number lise 1 - 97 and useing a foreach loop to go over every listing that tvguide has.
# I am then filtering the text on each listing only for the show I want to annoy my mom with in the general use case it will be Twilight.
# As I do not know how she does it but every time its on I'm at her house.

#To improve this run the foreach in parllel useing like 10 instances at the same time as its slow as shit.

# $shows = @(
#     'The Twlight Saga',
#     'The Shining',
#     'NOVA',
#     'Star Trek',
#     'The Conjuring')
$range = 1..97
Measure-Command{
foreach ($number in $range) {
    $Driver.FindElementByXPath("//*[@id='listings']/ul/li[$number]") | Export-Csv -Path "E:\TVGuideData\TVGideData.csv" -Append
}
}
# foreach ($number in $range) {
#     $Driver.FindElementByXPath("//*[@id='listings']/ul/li[$number]") | Where-Object -Property text -like "*$shows*" | Select-Object -ExpandProperty Text
# }

# $Driver.FindElementByXPath('//*[@id="listings"]/ul/li[17]') | Where-Object -Property text -like "*The Shining*" | Select-Object -ExpandProperty Text

Stop-Process -Name Firefox
