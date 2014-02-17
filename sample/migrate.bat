@echo Delete git directory
@if exist git rd git /s /q
@mkdir git

@echo Start miigration...
@pushd git
ruby ..\..\vss2git.rb -r1 -s ..\vss -u admin -c git -d mail.sample.co.jp -b2 $/ >..\migrate.log
@popd
