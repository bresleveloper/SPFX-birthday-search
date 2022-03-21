# spfx-birthdays-sp-search

uses sp search to get birthdays from user profile service

must make sure you make sync

Search Mapping:
* `People:SPS-Birthday` to `RefinableString99`
* `People:SPS-Department` to `RefinableString98` 
* `People:Department` to `RefinableString97`
* `People:SPS-DisabledUser` (if exist) to `RefinableString95`

if `RefinableString95` equals 0 then the user is disabled

`https://github.com/bresleveloper/SPFX-birthday-search`


## download sln and run

* `gulp build`
* `gulp bundle --ship`
* `gulp package-solution --ship`

gulp build; gulp bundle --ship; gulp package-solution --ship



npm ls -g --depth=0

