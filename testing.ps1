$script:newDepartmentAD = @("planning","it-team","newbuild","sales-team","technical-team","senior-managers","newbuild-management","NPD-project-planning")

"
Football
Boxing
Tennis
Basketball
Formula 1
Cricket
Baseball
Rugby Union
Ice Hockey
Am.Football
Rugby League
Athletics / Misc. Olympic events
Wrestling
Mixed Martial Arts
Cycling
Golf
Snooker / Pool
Horse racing
Skateboarding
Indy/Nascar
Darts

Lionel Messi
Diego Maradona
Cristiano Ronaldo
Pele

3 Sugar Ray Robinson
6 Muhammad Ali
Floyd Mayweather Jr.
Rocky Marciano
Manny Pacquiao

Rafael Nadal
Roger Federer
Margaret Court
Steffi Graf
Serena Williams
Novak Djokovic

1 Michael Jordan
Wilt Chamberlain
Kareem Abdul-Jabaar
LeBron James

Lewis Hamilton
Michael Schumacher
Ayrton Senna

5 Don Bradman
Shane Warne
Brian Lara
Sachin Tendulkar
James Anderson
Muttiah Muralidaran
Kumar Sangakkara

2 Babe Ruth
Willie Mays
Jackie Robinson
Hank Aaron
Barry Bonds

Richie McCaw
Dan Carter
Brian O'Driscoll
Gareth Edwards
Shane Williams
Jonah Lomu
Andrew Johns

4 Wayne Gretzky
Jaromir Jagr

7 Jim Thorpe
Bo Jackson
Tom Brady
Jim Brown
Jerry Rice

Jesse Owens
Nadia Comaneci
Simone Biles
Steve Redgrave
Carl Lewis
Michael Phelps
Paavo Nurmi
Haile Gebresellasie

Aleksandr Karelin
Anderson Silva
Jon Jones

Eddy Merckx
Pinault?
Chris Froome

8 Babe Didrikson
Tiger Woods
Jack Nicklaus

Ronnie O'Sullivan
Stephen Hendry

Secretariat
Citation
Red Rum

Rodney Mullen
Tony Hawk

Richard Petty

Phil Taylor
Raymond Van Barneveld

"

$k=0
for ($i=0; $i -lt 5; $i++)
{
    $script:getADuserTest = Get-ADuser -Identity "gcooper" -Properties MemberOf

    foreach ($department in $script:newDepartmentAD)
    {
        if
        (
            $script:getADuserTest.MemberOf -like "*$department*"
        )
        {"Found $department"; $k++}
            
        elseif ($i -lt 4)
            {$j=$i+1;
            Write-Error "Attempt ${j}/5. Unable to set access groups. Waiting for 10 seconds";
            Start-Sleep -Seconds 5.0;
            Start-Sleep -Seconds 5.0}
        else {Write-Error "Attempt 5/5. Unable to set access groups. Exiting"; return}

        if ($k -eq $script:newDepartmentAD.Count) {return}
    }
}