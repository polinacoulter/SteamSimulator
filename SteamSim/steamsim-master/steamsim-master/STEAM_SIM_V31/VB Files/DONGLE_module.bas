Attribute VB_Name = "DONGLE_module"

' === 11/09/03 ===
Global Super_Pro_Dongle As Boolean


'=== 03/02/03  ===
Global TRIAL_HAS_ENDED As Boolean

' === 30/10/02 Experimenting ===
Global SSPRO_start_date_address As Integer

Global System_current_date_INT As Integer

Global SSPRO_start_date As Integer
Global SSPRO_current_date As Integer

Global SSPRO_access_code As Integer

Global System_current_date_LONG As Long
Global Timeout_period As Integer
Global Write_count As Integer

Global Const SSPRO_current_date_address = 16
Global Const SSPRO_Write_Password = &H34B7
Global Const Timeout_period_address = 24

Global Const maxim100_MemAddress = 12
Global Const maxim200_MemAddress = 20
Global Const maxim300_MemAddress = 28
Global Const maxim400_MemAddress = 36
Global Const maxim500_MemAddress = 44
Global Const maxim600_MemAddress = 52

' === 03/02/03 ADDED: ===
Global Const maxim700_MemAddress = 60




'
'Declare Function OPEN_SENT Lib "Sp32w.dll" () As Integer
'Declare Function CLOSE_SENT Lib "Sp32w.dll" () As Integer
'Declare Function SSQUERY Lib "Sp32w.dll" (ByVal QueryStr$, ByVal querylen%) As Integer

'
'Global Dongle_Reading(200) As Integer
'===6/8/02 new dongle stuff===
Global Dongle_Reading(200) As String


Global Dongle_Response As DATAQUERY

Global UseDongleResponseStart As Integer
Global UseDongleResponseLength As Integer



Global QueryStr As DATAQUERY
Global ExpectedResponse As DATAQUERY
Global ResponseStr As DATAQUERY
Global unused As Long
Global QueryLength As Integer


'Global DongleResponse(200) As Integer



' === 15/05/02 ===
Global Dongle_2_Response(200) As Integer

Global DonglePassword(200) As String

Global Crash_it As Double
'
'



' (C) Copyright 1986-1997 Rainbow Technologies, Inc. All rights reserved.



'Type APIPACKET
' reserved As String * 1028
'End Type


Type APIPACKET
 data(4096) As Byte
End Type

Type DATAQUERY
 data(16) As Byte
End Type



Declare Function RNBOsproFormatPacket% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal ApiPackSize As Integer)
Declare Function RNBOsproInitialize% Lib "sx32w.dll" (ApiPack As APIPACKET)
Declare Function RNBOsproGetFullStatus% Lib "sx32w.dll" (ApiPack As APIPACKET)
Declare Function RNBOsproGetVersion% Lib "sx32w.dll" (ApiPack As APIPACKET, majv As Integer, minv As Integer, rev As Integer, ostype As Integer)
Declare Function RNBOsproFindFirstUnit% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal developerID As Integer)
Declare Function RNBOsproFindNextUnit% Lib "sx32w.dll" (ApiPack As APIPACKET)
Declare Function RNBOsproRead% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, datum As Integer)
Declare Function RNBOsproExtendedRead% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, datum As Integer, accessCode As Integer)
Declare Function RNBOsproWrite% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal address As Integer, ByVal datum As Integer, ByVal accessCode As Integer)
Declare Function RNBOsproOverwrite% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal oPass1 As Integer, ByVal oPass2 As Integer, ByVal address As Integer, ByVal datum As Integer, ByVal accessCode As Integer)
Declare Function RNBOsproDecrement% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal address As Integer)
Declare Function RNBOsproActivate% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal aPass1 As Integer, ByVal aPass2 As Integer, ByVal address As Integer)
Declare Function RNBOsproQuery% Lib "sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, Query As DATAQUERY, Response As DATAQUERY, unused As Long, ByVal length As Integer)

Global devID%
Global Adr%, AC%, datum%
Global q32Resp&, qLen%, qResp$, qStr$
Global aPW1%, aPW2%, owPW1%, owPW2%, wPW%
Global spStatus%
Global Result%
' =====(C) Copyright 1986-1997 Rainbow Technologies, Inc. All rights reserved. =====END

'===5/8/02 testing new dongle stuff===
Global MyAPIPacket As APIPACKET

Global DongleConnected As Boolean


Public Type TQueryPair
    Query As String
    Response As String
End Type

Global QueryTable(1 To 200) As TQueryPair

' === 11/09/03 ===
'Public Sub InitQueryTable_Max700()
Public Sub InitQueryTable_Max600()

QueryTable(1).Query = "9C23"
QueryTable(1).Response = "401C"
QueryTable(2).Query = "C6A6"
QueryTable(2).Response = "861C"
QueryTable(3).Query = "5B30"
QueryTable(3).Response = "EC84"
QueryTable(4).Query = "08E9"
QueryTable(4).Response = "43B3"
QueryTable(5).Query = "C0B8"
QueryTable(5).Response = "9C32"
QueryTable(6).Query = "876B"
QueryTable(6).Response = "BEB4"
QueryTable(7).Query = "7CFE"
QueryTable(7).Response = "527A"
QueryTable(8).Query = "8733"
QueryTable(8).Response = "1FF1"
QueryTable(9).Query = "0BDC"
QueryTable(9).Response = "1B03"
QueryTable(10).Query = "8B87"
QueryTable(10).Response = "ADBF"
QueryTable(11).Query = "2C66"
QueryTable(11).Response = "703B"
QueryTable(12).Query = "6CAF"
QueryTable(12).Response = "0414"
QueryTable(13).Query = "94CA"
QueryTable(13).Response = "E7E2"
QueryTable(14).Query = "A9FC"
QueryTable(14).Response = "C04A"
QueryTable(15).Query = "E37A"
QueryTable(15).Response = "21B3"
QueryTable(16).Query = "FE09"
QueryTable(16).Response = "6D65"
QueryTable(17).Query = "D78E"
QueryTable(17).Response = "0AB1"
QueryTable(18).Query = "3841"
QueryTable(18).Response = "FBB8"
QueryTable(19).Query = "C063"
QueryTable(19).Response = "AD49"
QueryTable(20).Query = "47A1"
QueryTable(20).Response = "A46B"
QueryTable(21).Query = "348995"
QueryTable(21).Response = "012467"
QueryTable(22).Query = "BF03DC"
QueryTable(22).Response = "CAECA5"
QueryTable(23).Query = "960C6D"
QueryTable(23).Response = "3CF9A1"
QueryTable(24).Query = "F0A056"
QueryTable(24).Response = "E288ED"
QueryTable(25).Query = "93B3FE"
QueryTable(25).Response = "18A0C0"
QueryTable(26).Query = "912C4D"
QueryTable(26).Response = "D95BA4"
QueryTable(27).Query = "3FD679"
QueryTable(27).Response = "BFA677"
QueryTable(28).Query = "4BD4AC"
QueryTable(28).Response = "98C727"
QueryTable(29).Query = "7BFDDB"
QueryTable(29).Response = "9B55C0"
QueryTable(30).Query = "87631C"
QueryTable(30).Response = "F6D205"
QueryTable(31).Query = "358E4A"
QueryTable(31).Response = "87E945"
QueryTable(32).Query = "B0DEA1"
QueryTable(32).Response = "5536F0"
QueryTable(33).Query = "22317E"
QueryTable(33).Response = "234697"
QueryTable(34).Query = "9BEDED"
QueryTable(34).Response = "4CEAE5"
QueryTable(35).Query = "3BD42C"
QueryTable(35).Response = "EE87D5"
QueryTable(36).Query = "A9D2EA"
QueryTable(36).Response = "D66FC8"
QueryTable(37).Query = "9B8BE6"
QueryTable(37).Response = "0328F1"
QueryTable(38).Query = "C21B38"
QueryTable(38).Response = "88B49F"
QueryTable(39).Query = "97C9F0"
QueryTable(39).Response = "B112C7"
QueryTable(40).Query = "395780"
QueryTable(40).Response = "8EE587"
QueryTable(41).Query = "2CA3F937"
QueryTable(41).Response = "CA7D11A0"
QueryTable(42).Query = "6BF70B18"
QueryTable(42).Response = "BB6473C3"
QueryTable(43).Query = "C1476668"
QueryTable(43).Response = "2D334FCD"
QueryTable(44).Query = "F19A030A"
QueryTable(44).Response = "9D9F5782"
QueryTable(45).Query = "E0921145"
QueryTable(45).Response = "9067AF3B"
QueryTable(46).Query = "494C44C6"
QueryTable(46).Response = "8364FBA2"
QueryTable(47).Query = "B1886FE9"
QueryTable(47).Response = "E334D860"
QueryTable(48).Query = "9E0A5534"
QueryTable(48).Response = "155B4C23"
QueryTable(49).Query = "043FBC24"
QueryTable(49).Response = "FD840A42"
QueryTable(50).Query = "0024C42D"
QueryTable(50).Response = "FBD4BA90"
QueryTable(51).Query = "C86477FC"
QueryTable(51).Response = "B6EBA154"
QueryTable(52).Query = "E2C0A1FF"
QueryTable(52).Response = "53FB1FC4"
QueryTable(53).Query = "95B4E227"
QueryTable(53).Response = "8E704090"
QueryTable(54).Query = "9E5801ED"
QueryTable(54).Response = "0CB36F35"
QueryTable(55).Query = "23848393"
QueryTable(55).Response = "ACAF09D2"
QueryTable(56).Query = "ED377DAC"
QueryTable(56).Response = "E53653B3"
QueryTable(57).Query = "D13AA6DE"
QueryTable(57).Response = "FF7176EE"
QueryTable(58).Query = "6F00B0E7"
QueryTable(58).Response = "6CA69560"
QueryTable(59).Query = "1FD1D9DE"
QueryTable(59).Response = "94E3AD8B"
QueryTable(60).Query = "2927BDBB"
QueryTable(60).Response = "5A698F81"
QueryTable(61).Query = "34596E1A02"
QueryTable(61).Response = "443D72AE6A"
QueryTable(62).Query = "76C63C5A74"
QueryTable(62).Response = "07EFF269CB"
QueryTable(63).Query = "FC4F458A76"
QueryTable(63).Response = "0997688C48"
QueryTable(64).Query = "EF7CDADEEC"
QueryTable(64).Response = "F01719A83F"
QueryTable(65).Query = "2253784EAF"
QueryTable(65).Response = "0FE86A6E88"
QueryTable(66).Query = "D6B3962C97"
QueryTable(66).Response = "F168CC1C28"
QueryTable(67).Query = "FD47A6C7DC"
QueryTable(67).Response = "EEB5960648"
QueryTable(68).Query = "A5DD703BBD"
QueryTable(68).Response = "0BC1634167"
QueryTable(69).Query = "3CA001A50E"
QueryTable(69).Response = "EDFD6ECF97"
QueryTable(70).Query = "C9E905EE68"
QueryTable(70).Response = "2AE670601C"
QueryTable(71).Query = "8DD67FE5C3"
QueryTable(71).Response = "4DD6B4C792"
QueryTable(72).Query = "6A5EBA2DAD"
QueryTable(72).Response = "B99D705AA2"
QueryTable(73).Query = "55D80CAAD2"
QueryTable(73).Response = "0FC9BDBBFE"
QueryTable(74).Query = "465551970C"
QueryTable(74).Response = "9D589FBC9A"
QueryTable(75).Query = "0777B24FBB"
QueryTable(75).Response = "FA1A06116B"
QueryTable(76).Query = "1EB0BD7DF7"
QueryTable(76).Response = "D1F54AEBC2"
QueryTable(77).Query = "0E2BDA25F0"
QueryTable(77).Response = "E17866D9D1"
QueryTable(78).Query = "7D3A4415C7"
QueryTable(78).Response = "C91F1868A6"
QueryTable(79).Query = "70AA0BB190"
QueryTable(79).Response = "AD52388B64"
QueryTable(80).Query = "EC9F2001A4"
QueryTable(80).Response = "1EE3ECB242"
QueryTable(81).Query = "9813D372A5CA"
QueryTable(81).Response = "D9EC0EBA78BD"
QueryTable(82).Query = "5241932EDBCC"
QueryTable(82).Response = "BC038F032E2D"
QueryTable(83).Query = "C36DEE60DCE5"
QueryTable(83).Response = "B050B69C3FF5"
QueryTable(84).Query = "85E3BD43A415"
QueryTable(84).Response = "30B1ECCC569C"
QueryTable(85).Query = "B5EC3216F79B"
QueryTable(85).Response = "3D2093FE3FC8"
QueryTable(86).Query = "0728484C46B7"
QueryTable(86).Response = "041605D07C3C"
QueryTable(87).Query = "FB5AED961FE4"
QueryTable(87).Response = "9244911B48D3"
QueryTable(88).Query = "AA8CF774F00F"
QueryTable(88).Response = "0A2DE203054E"
QueryTable(89).Query = "AAE34B4F3C77"
QueryTable(89).Response = "988C2C3BC7CE"
QueryTable(90).Query = "F177BD4F992F"
QueryTable(90).Response = "DB3FE7D9C5D6"
QueryTable(91).Query = "C10190650F81"
QueryTable(91).Response = "F362676D36CC"
QueryTable(92).Query = "094D9E45B4A6"
QueryTable(92).Response = "667DCA7A2756"
QueryTable(93).Query = "BDA6933F918A"
QueryTable(93).Response = "5641B4DF9867"
QueryTable(94).Query = "09BBC4403491"
QueryTable(94).Response = "38B25C678B91"
QueryTable(95).Query = "7CAB7D568AA3"
QueryTable(95).Response = "52F586D6EAD7"
QueryTable(96).Query = "6E3FE498C3E2"
QueryTable(96).Response = "22FB60C17512"
QueryTable(97).Query = "5678C33C74D3"
QueryTable(97).Response = "31763FAD8A23"
QueryTable(98).Query = "D1F8D33D3EEE"
QueryTable(98).Response = "1B2A1436712B"
QueryTable(99).Query = "79F561EB91C3"
QueryTable(99).Response = "54D9A9EDCC98"
QueryTable(100).Query = "EAB06327733D"
QueryTable(100).Response = "F6D90545B88C"
QueryTable(101).Query = "84AF584F4AA9C9"
QueryTable(101).Response = "A752E6D8C0BDE3"
QueryTable(102).Query = "2783FA1D8D674E"
QueryTable(102).Response = "6EA7BA59365125"
QueryTable(103).Query = "6915CA7252A196"
QueryTable(103).Response = "5F6A1AB9369328"
QueryTable(104).Query = "FD1F27793612B4"
QueryTable(104).Response = "836F9324171575"
QueryTable(105).Query = "DB32F9B1133D00"
QueryTable(105).Response = "B957194D1D14FB"
QueryTable(106).Query = "65DA7BC9BC00CF"
QueryTable(106).Response = "5F94CB1244199D"
QueryTable(107).Query = "BF475DFCA7830B"
QueryTable(107).Response = "E6BF5E045FCF22"
QueryTable(108).Query = "0B775ED2EC6CF7"
QueryTable(108).Response = "B808ACAC708ED6"
QueryTable(109).Query = "E45F57A6CB1537"
QueryTable(109).Response = "A5BE87EE2B6B62"
QueryTable(110).Query = "3C9F130D8C7C58"
QueryTable(110).Response = "4F4F7B145A27C0"
QueryTable(111).Query = "EC2A33BF507240"
QueryTable(111).Response = "E8A0FBBC4420EF"
QueryTable(112).Query = "27B0CCFDD0572F"
QueryTable(112).Response = "AE07CD6DCCA1E8"
QueryTable(113).Query = "14086FE43A085F"
QueryTable(113).Response = "E5C2DBED508FFB"
QueryTable(114).Query = "8D1B19576590FA"
QueryTable(114).Response = "5E7186D319DB30"
QueryTable(115).Query = "28C918865AF886"
QueryTable(115).Response = "50E263FABFDF85"
QueryTable(116).Query = "927817F715AF77"
QueryTable(116).Response = "07D8A34E0E10B3"
QueryTable(117).Query = "078B971B459B51"
QueryTable(117).Response = "7BE4BF196716BA"
QueryTable(118).Query = "75E3F7D271B283"
QueryTable(118).Response = "7CE7B66D719115"
QueryTable(119).Query = "7EB44EBAF92696"
QueryTable(119).Response = "3B5431E1FDFF0A"
QueryTable(120).Query = "4E3AAE600FE13C"
QueryTable(120).Response = "2C1B086D9C20A1"
QueryTable(121).Query = "6A0ECA0D9302E22A"
QueryTable(121).Response = "EF0B51E926BDC826"
QueryTable(122).Query = "4BEB393348F5595A"
QueryTable(122).Response = "DBDEFEF263E904EA"
QueryTable(123).Query = "7A9ED6B2421B27AA"
QueryTable(123).Response = "9395251AAA256309"
QueryTable(124).Query = "DFF8B4D4280030EC"
QueryTable(124).Response = "3AADB6C19CB84555"
QueryTable(125).Query = "815C291508FEDE65"
QueryTable(125).Response = "8CCC138619390653"
QueryTable(126).Query = "0F4150A4CFC99C51"
QueryTable(126).Response = "115A0D10B293AB56"
QueryTable(127).Query = "5AB187E6D2E55B5D"
QueryTable(127).Response = "D763EB21B8970973"
QueryTable(128).Query = "D9CCF2F3482B102C"
QueryTable(128).Response = "CB61D9259C95AAA6"
QueryTable(129).Query = "2446F517CD4733D3"
QueryTable(129).Response = "09BB22A7D1974E8B"
QueryTable(130).Query = "79E4BE52E13D425D"
QueryTable(130).Response = "1F0862390F09D999"
QueryTable(131).Query = "3A03B9D967DF3C48"
QueryTable(131).Response = "104C3D04400ACF71"
QueryTable(132).Query = "6D12199227582706"
QueryTable(132).Response = "175C73CC445079C5"
QueryTable(133).Query = "3B13569C4CA58A7D"
QueryTable(133).Response = "816F5A00AB1928DE"
QueryTable(134).Query = "721EAAC5E717F387"
QueryTable(134).Response = "8D389ED3D4190895"
QueryTable(135).Query = "02DD93116AD37372"
QueryTable(135).Response = "A1BE4CE6DC634AB0"
QueryTable(136).Query = "8311543A2D511D81"
QueryTable(136).Response = "A0B4A775F39D6B98"
QueryTable(137).Query = "AE0C732BECE08B69"
QueryTable(137).Response = "A6BEDA9DA586CBAD"
QueryTable(138).Query = "E0353C83472059D6"
QueryTable(138).Response = "E62923E31CA993FC"
QueryTable(139).Query = "9B873D184187A7E4"
QueryTable(139).Response = "4DD010B5F1E0530E"
QueryTable(140).Query = "0412C872C2DD9BA5"
QueryTable(140).Response = "6E14C264336CC42B"
QueryTable(141).Query = "6579754C15BFDDA1AC"
QueryTable(141).Response = "CE1C600C98194A263F"
QueryTable(142).Query = "729D166B201952EB49"
QueryTable(142).Response = "5878F0E37B2AE42451"
QueryTable(143).Query = "E07656C381A4D65DA0"
QueryTable(143).Response = "D8683161A209B7F8E5"
QueryTable(144).Query = "C24EC3497B49A28387"
QueryTable(144).Response = "382AA1FA0471168209"
QueryTable(145).Query = "2F0C292DBF1DF506B9"
QueryTable(145).Response = "7BE26B2DAA223DDBDA"
QueryTable(146).Query = "DFDE03DB6BA4B10F51"
QueryTable(146).Response = "127B7506AC5AECB9F8"
QueryTable(147).Query = "A9BCE33A02B239CCCE"
QueryTable(147).Response = "BD2AFB97A9834F2D1A"
QueryTable(148).Query = "0A40CEC665A38D1613"
QueryTable(148).Response = "463C8265154AC8D67B"
QueryTable(149).Query = "007E6BD99B244EF254"
QueryTable(149).Response = "222156A9FCBF44F413"
QueryTable(150).Query = "34B053D6F8C63312F1"
QueryTable(150).Response = "9DA9B7C91DF1E6D5C6"
QueryTable(151).Query = "17CBEBC0D25B334666"
QueryTable(151).Response = "CEC1F57CF010BEA071"
QueryTable(152).Query = "6DD70B4D5D56DC2819"
QueryTable(152).Response = "71438C311F629B05BC"
QueryTable(153).Query = "78B811263CEE2DE412"
QueryTable(153).Response = "670A4D5E14372872C8"
QueryTable(154).Query = "7F80124AC90694CD67"
QueryTable(154).Response = "31A3927B29F4DA5AB5"
QueryTable(155).Query = "FF740B538A28344D75"
QueryTable(155).Response = "9095C3C77CEEFC2286"
QueryTable(156).Query = "13AB81544DB764D0A7"
QueryTable(156).Response = "772BA5ADB40174FBCB"
QueryTable(157).Query = "7CDF42EA2BAF806075"
QueryTable(157).Response = "18D20FA5247692E3D4"
QueryTable(158).Query = "0BF7964588FB5FAE3B"
QueryTable(158).Response = "A7D8A03929961F0843"
QueryTable(159).Query = "5A1C221D373BE6B316"
QueryTable(159).Response = "21F1985B87ED8E2D64"
QueryTable(160).Query = "FEC2EE337EFD4F18EF"
QueryTable(160).Response = "A4363E3893866588A7"
QueryTable(161).Query = "88506E257A001457BCED"
QueryTable(161).Response = "F041ABA3036F4473CCB5"
QueryTable(162).Query = "6E24064600D30A413BDF"
QueryTable(162).Response = "E416FA42973ED841D7B6"
QueryTable(163).Query = "EEBB96B3001EE1B2A688"
QueryTable(163).Response = "1F08A140C7D01F692D67"
QueryTable(164).Query = "D67A932AF26A41ED8261"
QueryTable(164).Response = "E130762AAA138CF63BFB"
QueryTable(165).Query = "9B5C01C9F9BEA1D55282"
QueryTable(165).Response = "FA7F0E03519FB28CA098"
QueryTable(166).Query = "2806ADE1FDF4D9BE834B"
QueryTable(166).Response = "F4F550949107BEDE7CE1"
QueryTable(167).Query = "5ED1A514B304B236B21D"
QueryTable(167).Response = "400A920B0B89FBFD185D"
QueryTable(168).Query = "6C3C489E5B5871E968A8"
QueryTable(168).Response = "4DED5ECB4ADCBBD5E503"
QueryTable(169).Query = "64A8A6B90A6C734AB594"
QueryTable(169).Response = "0C2E4B15077C47105ADB"
QueryTable(170).Query = "C0520F71AD4350A75991"
QueryTable(170).Response = "9DA61CBF4F5A9479A5F2"
QueryTable(171).Query = "B5D7F86E2D7B26D85511"
QueryTable(171).Response = "EEB86461DD2D92F291E7"
QueryTable(172).Query = "7AD6A67B5FFFEDA61F3B"
QueryTable(172).Response = "FF70A805ACB9D11B801E"
QueryTable(173).Query = "FC6A41CEA0B11592D3DC"
QueryTable(173).Response = "1722CDAFE2644BFB554A"
QueryTable(174).Query = "8D8332624E81B1BA0752"
QueryTable(174).Response = "A9708812F9473BCD8800"
QueryTable(175).Query = "8D7C02E782C3F8E03ED5"
QueryTable(175).Response = "CD6AF97BDEEC2B374642"
QueryTable(176).Query = "2A511606ADC7F9D90686"
QueryTable(176).Response = "5B53A7AB936AD99CE38E"
QueryTable(177).Query = "A269F615110797FD3B25"
QueryTable(177).Response = "AB8D97464B0D2F3D3044"
QueryTable(178).Query = "A8CD017226615F471C64"
QueryTable(178).Response = "6302D1DFFFA08908C254"
QueryTable(179).Query = "E642BB146C30AD450231"
QueryTable(179).Response = "895DA5150DF6B3F1246E"
QueryTable(180).Query = "A7A926163E4E3A0EFB68"
QueryTable(180).Response = "A9B78D863EC7C9BB9F35"
QueryTable(181).Query = "3582B151993D0ECEA4D071"
QueryTable(181).Response = "B6DFD3D708824814852065"
QueryTable(182).Query = "94CB35FB0B64B2E951A300"
QueryTable(182).Response = "685450F2EB73326CE37A69"
QueryTable(183).Query = "2182C7AD0A3795BFA14C06"
QueryTable(183).Response = "7BB3B62EE658AB11B4004F"
QueryTable(184).Query = "9ACAE73D24E3B5BB269F77"
QueryTable(184).Response = "FE44CFABB9DA087EC5CD11"
QueryTable(185).Query = "D6F512BE675308F2C4909B"
QueryTable(185).Response = "8A9AD40C7A446EA6D87420"
QueryTable(186).Query = "99EEF405D60B61B2301635"
QueryTable(186).Response = "239C07F102EF130FBDE6BE"
QueryTable(187).Query = "9995D56ECD9F0B114F15EF"
QueryTable(187).Response = "00BF01692D73125EFE70E3"
QueryTable(188).Query = "CC27B5DEDAB3314D273E77"
QueryTable(188).Response = "50D5556775A6BC1900618D"
QueryTable(189).Query = "01B5D435823E559DCD1BC9"
QueryTable(189).Response = "C6423CA273BFFB6F6B11E1"
QueryTable(190).Query = "654964F313FC592977FF00"
QueryTable(190).Response = "D88022150F8ACD8B717B81"
QueryTable(191).Query = "62792F109114260E32A369"
QueryTable(191).Response = "797FB1840DEA2B069404C9"
QueryTable(192).Query = "B6225CE39CB31A9313005F"
QueryTable(192).Response = "CC3F1E743066D36032C9F7"
QueryTable(193).Query = "814E7E99F6C5D4A4ECDAED"
QueryTable(193).Response = "48024C895432004F5846C5"
QueryTable(194).Query = "91C4BD4BA6378D9EA58FC0"
QueryTable(194).Response = "7B0D20031B8D2F5543B728"
QueryTable(195).Query = "955117C4F42FD0E6E68A65"
QueryTable(195).Response = "BF0863D6D7DBBDA2F96D31"
QueryTable(196).Query = "7495B9611EF9A4A0A2ED85"
QueryTable(196).Response = "DAAB392D7E57D9B1E7D4E9"
QueryTable(197).Query = "03A988EC68EBAA06DA9AD5"
QueryTable(197).Response = "64965E421548FAFB5465E5"
QueryTable(198).Query = "1E04FE3F64EC5C25EEDE1D"
QueryTable(198).Response = "6CC8B47867B45D23C58F8B"
QueryTable(199).Query = "61FCA52C3F8001A08DDE56"
QueryTable(199).Response = "C2A9CED8C3567FBD0CAC7F"
QueryTable(200).Query = "9BF1F0A2727F4DB10E8BCB"
QueryTable(200).Response = "DB32839DE416BFE4ED5E28"
QueryTable(1).Query = "9C23"
QueryTable(1).Response = "401C"
QueryTable(2).Query = "C6A6"
QueryTable(2).Response = "861C"
QueryTable(3).Query = "5B30"
QueryTable(3).Response = "EC84"
QueryTable(4).Query = "08E9"
QueryTable(4).Response = "43B3"
QueryTable(5).Query = "C0B8"
QueryTable(5).Response = "9C32"
QueryTable(6).Query = "876B"
QueryTable(6).Response = "BEB4"
QueryTable(7).Query = "7CFE"
QueryTable(7).Response = "527A"
QueryTable(8).Query = "8733"
QueryTable(8).Response = "1FF1"
QueryTable(9).Query = "0BDC"
QueryTable(9).Response = "1B03"
QueryTable(10).Query = "8B87"
QueryTable(10).Response = "ADBF"
QueryTable(11).Query = "2C66"
QueryTable(11).Response = "703B"
QueryTable(12).Query = "6CAF"
QueryTable(12).Response = "0414"
QueryTable(13).Query = "94CA"
QueryTable(13).Response = "E7E2"
QueryTable(14).Query = "A9FC"
QueryTable(14).Response = "C04A"
QueryTable(15).Query = "E37A"
QueryTable(15).Response = "21B3"
QueryTable(16).Query = "FE09"
QueryTable(16).Response = "6D65"
QueryTable(17).Query = "D78E"
QueryTable(17).Response = "0AB1"
QueryTable(18).Query = "3841"
QueryTable(18).Response = "FBB8"
QueryTable(19).Query = "C063"
QueryTable(19).Response = "AD49"
QueryTable(20).Query = "47A1"
QueryTable(20).Response = "A46B"
QueryTable(21).Query = "348995"
QueryTable(21).Response = "012467"
QueryTable(22).Query = "BF03DC"
QueryTable(22).Response = "CAECA5"
QueryTable(23).Query = "960C6D"
QueryTable(23).Response = "3CF9A1"
QueryTable(24).Query = "F0A056"
QueryTable(24).Response = "E288ED"
QueryTable(25).Query = "93B3FE"
QueryTable(25).Response = "18A0C0"
QueryTable(26).Query = "912C4D"
QueryTable(26).Response = "D95BA4"
QueryTable(27).Query = "3FD679"
QueryTable(27).Response = "BFA677"
QueryTable(28).Query = "4BD4AC"
QueryTable(28).Response = "98C727"
QueryTable(29).Query = "7BFDDB"
QueryTable(29).Response = "9B55C0"
QueryTable(30).Query = "87631C"
QueryTable(30).Response = "F6D205"
QueryTable(31).Query = "358E4A"
QueryTable(31).Response = "87E945"
QueryTable(32).Query = "B0DEA1"
QueryTable(32).Response = "5536F0"
QueryTable(33).Query = "22317E"
QueryTable(33).Response = "234697"
QueryTable(34).Query = "9BEDED"
QueryTable(34).Response = "4CEAE5"
QueryTable(35).Query = "3BD42C"
QueryTable(35).Response = "EE87D5"
QueryTable(36).Query = "A9D2EA"
QueryTable(36).Response = "D66FC8"
QueryTable(37).Query = "9B8BE6"
QueryTable(37).Response = "0328F1"
QueryTable(38).Query = "C21B38"
QueryTable(38).Response = "88B49F"
QueryTable(39).Query = "97C9F0"
QueryTable(39).Response = "B112C7"
QueryTable(40).Query = "395780"
QueryTable(40).Response = "8EE587"
QueryTable(41).Query = "2CA3F937"
QueryTable(41).Response = "CA7D11A0"
QueryTable(42).Query = "6BF70B18"
QueryTable(42).Response = "BB6473C3"
QueryTable(43).Query = "C1476668"
QueryTable(43).Response = "2D334FCD"
QueryTable(44).Query = "F19A030A"
QueryTable(44).Response = "9D9F5782"
QueryTable(45).Query = "E0921145"
QueryTable(45).Response = "9067AF3B"
QueryTable(46).Query = "494C44C6"
QueryTable(46).Response = "8364FBA2"
QueryTable(47).Query = "B1886FE9"
QueryTable(47).Response = "E334D860"
QueryTable(48).Query = "9E0A5534"
QueryTable(48).Response = "155B4C23"
QueryTable(49).Query = "043FBC24"
QueryTable(49).Response = "FD840A42"
QueryTable(50).Query = "0024C42D"
QueryTable(50).Response = "FBD4BA90"
QueryTable(51).Query = "C86477FC"
QueryTable(51).Response = "B6EBA154"
QueryTable(52).Query = "E2C0A1FF"
QueryTable(52).Response = "53FB1FC4"
QueryTable(53).Query = "95B4E227"
QueryTable(53).Response = "8E704090"
QueryTable(54).Query = "9E5801ED"
QueryTable(54).Response = "0CB36F35"
QueryTable(55).Query = "23848393"
QueryTable(55).Response = "ACAF09D2"
QueryTable(56).Query = "ED377DAC"
QueryTable(56).Response = "E53653B3"
QueryTable(57).Query = "D13AA6DE"
QueryTable(57).Response = "FF7176EE"
QueryTable(58).Query = "6F00B0E7"
QueryTable(58).Response = "6CA69560"
QueryTable(59).Query = "1FD1D9DE"
QueryTable(59).Response = "94E3AD8B"
QueryTable(60).Query = "2927BDBB"
QueryTable(60).Response = "5A698F81"
QueryTable(61).Query = "34596E1A02"
QueryTable(61).Response = "443D72AE6A"
QueryTable(62).Query = "76C63C5A74"
QueryTable(62).Response = "07EFF269CB"
QueryTable(63).Query = "FC4F458A76"
QueryTable(63).Response = "0997688C48"
QueryTable(64).Query = "EF7CDADEEC"
QueryTable(64).Response = "F01719A83F"
QueryTable(65).Query = "2253784EAF"
QueryTable(65).Response = "0FE86A6E88"
QueryTable(66).Query = "D6B3962C97"
QueryTable(66).Response = "F168CC1C28"
QueryTable(67).Query = "FD47A6C7DC"
QueryTable(67).Response = "EEB5960648"
QueryTable(68).Query = "A5DD703BBD"
QueryTable(68).Response = "0BC1634167"
QueryTable(69).Query = "3CA001A50E"
QueryTable(69).Response = "EDFD6ECF97"
QueryTable(70).Query = "C9E905EE68"
QueryTable(70).Response = "2AE670601C"
QueryTable(71).Query = "8DD67FE5C3"
QueryTable(71).Response = "4DD6B4C792"
QueryTable(72).Query = "6A5EBA2DAD"
QueryTable(72).Response = "B99D705AA2"
QueryTable(73).Query = "55D80CAAD2"
QueryTable(73).Response = "0FC9BDBBFE"
QueryTable(74).Query = "465551970C"
QueryTable(74).Response = "9D589FBC9A"
QueryTable(75).Query = "0777B24FBB"
QueryTable(75).Response = "FA1A06116B"
QueryTable(76).Query = "1EB0BD7DF7"
QueryTable(76).Response = "D1F54AEBC2"
QueryTable(77).Query = "0E2BDA25F0"
QueryTable(77).Response = "E17866D9D1"
QueryTable(78).Query = "7D3A4415C7"
QueryTable(78).Response = "C91F1868A6"
QueryTable(79).Query = "70AA0BB190"
QueryTable(79).Response = "AD52388B64"
QueryTable(80).Query = "EC9F2001A4"
QueryTable(80).Response = "1EE3ECB242"
QueryTable(81).Query = "9813D372A5CA"
QueryTable(81).Response = "D9EC0EBA78BD"
QueryTable(82).Query = "5241932EDBCC"
QueryTable(82).Response = "BC038F032E2D"
QueryTable(83).Query = "C36DEE60DCE5"
QueryTable(83).Response = "B050B69C3FF5"
QueryTable(84).Query = "85E3BD43A415"
QueryTable(84).Response = "30B1ECCC569C"
QueryTable(85).Query = "B5EC3216F79B"
QueryTable(85).Response = "3D2093FE3FC8"
QueryTable(86).Query = "0728484C46B7"
QueryTable(86).Response = "041605D07C3C"
QueryTable(87).Query = "FB5AED961FE4"
QueryTable(87).Response = "9244911B48D3"
QueryTable(88).Query = "AA8CF774F00F"
QueryTable(88).Response = "0A2DE203054E"
QueryTable(89).Query = "AAE34B4F3C77"
QueryTable(89).Response = "988C2C3BC7CE"
QueryTable(90).Query = "F177BD4F992F"
QueryTable(90).Response = "DB3FE7D9C5D6"
QueryTable(91).Query = "C10190650F81"
QueryTable(91).Response = "F362676D36CC"
QueryTable(92).Query = "094D9E45B4A6"
QueryTable(92).Response = "667DCA7A2756"
QueryTable(93).Query = "BDA6933F918A"
QueryTable(93).Response = "5641B4DF9867"
QueryTable(94).Query = "09BBC4403491"
QueryTable(94).Response = "38B25C678B91"
QueryTable(95).Query = "7CAB7D568AA3"
QueryTable(95).Response = "52F586D6EAD7"
QueryTable(96).Query = "6E3FE498C3E2"
QueryTable(96).Response = "22FB60C17512"
QueryTable(97).Query = "5678C33C74D3"
QueryTable(97).Response = "31763FAD8A23"
QueryTable(98).Query = "D1F8D33D3EEE"
QueryTable(98).Response = "1B2A1436712B"
QueryTable(99).Query = "79F561EB91C3"
QueryTable(99).Response = "54D9A9EDCC98"
QueryTable(100).Query = "EAB06327733D"
QueryTable(100).Response = "F6D90545B88C"
QueryTable(101).Query = "84AF584F4AA9C9"
QueryTable(101).Response = "A752E6D8C0BDE3"
QueryTable(102).Query = "2783FA1D8D674E"
QueryTable(102).Response = "6EA7BA59365125"
QueryTable(103).Query = "6915CA7252A196"
QueryTable(103).Response = "5F6A1AB9369328"
QueryTable(104).Query = "FD1F27793612B4"
QueryTable(104).Response = "836F9324171575"
QueryTable(105).Query = "DB32F9B1133D00"
QueryTable(105).Response = "B957194D1D14FB"
QueryTable(106).Query = "65DA7BC9BC00CF"
QueryTable(106).Response = "5F94CB1244199D"
QueryTable(107).Query = "BF475DFCA7830B"
QueryTable(107).Response = "E6BF5E045FCF22"
QueryTable(108).Query = "0B775ED2EC6CF7"
QueryTable(108).Response = "B808ACAC708ED6"
QueryTable(109).Query = "E45F57A6CB1537"
QueryTable(109).Response = "A5BE87EE2B6B62"
QueryTable(110).Query = "3C9F130D8C7C58"
QueryTable(110).Response = "4F4F7B145A27C0"
QueryTable(111).Query = "EC2A33BF507240"
QueryTable(111).Response = "E8A0FBBC4420EF"
QueryTable(112).Query = "27B0CCFDD0572F"
QueryTable(112).Response = "AE07CD6DCCA1E8"
QueryTable(113).Query = "14086FE43A085F"
QueryTable(113).Response = "E5C2DBED508FFB"
QueryTable(114).Query = "8D1B19576590FA"
QueryTable(114).Response = "5E7186D319DB30"
QueryTable(115).Query = "28C918865AF886"
QueryTable(115).Response = "50E263FABFDF85"
QueryTable(116).Query = "927817F715AF77"
QueryTable(116).Response = "07D8A34E0E10B3"
QueryTable(117).Query = "078B971B459B51"
QueryTable(117).Response = "7BE4BF196716BA"
QueryTable(118).Query = "75E3F7D271B283"
QueryTable(118).Response = "7CE7B66D719115"
QueryTable(119).Query = "7EB44EBAF92696"
QueryTable(119).Response = "3B5431E1FDFF0A"
QueryTable(120).Query = "4E3AAE600FE13C"
QueryTable(120).Response = "2C1B086D9C20A1"
QueryTable(121).Query = "6A0ECA0D9302E22A"
QueryTable(121).Response = "EF0B51E926BDC826"
QueryTable(122).Query = "4BEB393348F5595A"
QueryTable(122).Response = "DBDEFEF263E904EA"
QueryTable(123).Query = "7A9ED6B2421B27AA"
QueryTable(123).Response = "9395251AAA256309"
QueryTable(124).Query = "DFF8B4D4280030EC"
QueryTable(124).Response = "3AADB6C19CB84555"
QueryTable(125).Query = "815C291508FEDE65"
QueryTable(125).Response = "8CCC138619390653"
QueryTable(126).Query = "0F4150A4CFC99C51"
QueryTable(126).Response = "115A0D10B293AB56"
QueryTable(127).Query = "5AB187E6D2E55B5D"
QueryTable(127).Response = "D763EB21B8970973"
QueryTable(128).Query = "D9CCF2F3482B102C"
QueryTable(128).Response = "CB61D9259C95AAA6"
QueryTable(129).Query = "2446F517CD4733D3"
QueryTable(129).Response = "09BB22A7D1974E8B"
QueryTable(130).Query = "79E4BE52E13D425D"
QueryTable(130).Response = "1F0862390F09D999"
QueryTable(131).Query = "3A03B9D967DF3C48"
QueryTable(131).Response = "104C3D04400ACF71"
QueryTable(132).Query = "6D12199227582706"
QueryTable(132).Response = "175C73CC445079C5"
QueryTable(133).Query = "3B13569C4CA58A7D"
QueryTable(133).Response = "816F5A00AB1928DE"
QueryTable(134).Query = "721EAAC5E717F387"
QueryTable(134).Response = "8D389ED3D4190895"
QueryTable(135).Query = "02DD93116AD37372"
QueryTable(135).Response = "A1BE4CE6DC634AB0"
QueryTable(136).Query = "8311543A2D511D81"
QueryTable(136).Response = "A0B4A775F39D6B98"
QueryTable(137).Query = "AE0C732BECE08B69"
QueryTable(137).Response = "A6BEDA9DA586CBAD"
QueryTable(138).Query = "E0353C83472059D6"
QueryTable(138).Response = "E62923E31CA993FC"
QueryTable(139).Query = "9B873D184187A7E4"
QueryTable(139).Response = "4DD010B5F1E0530E"
QueryTable(140).Query = "0412C872C2DD9BA5"
QueryTable(140).Response = "6E14C264336CC42B"
QueryTable(141).Query = "6579754C15BFDDA1AC"
QueryTable(141).Response = "CE1C600C98194A263F"
QueryTable(142).Query = "729D166B201952EB49"
QueryTable(142).Response = "5878F0E37B2AE42451"
QueryTable(143).Query = "E07656C381A4D65DA0"
QueryTable(143).Response = "D8683161A209B7F8E5"
QueryTable(144).Query = "C24EC3497B49A28387"
QueryTable(144).Response = "382AA1FA0471168209"
QueryTable(145).Query = "2F0C292DBF1DF506B9"
QueryTable(145).Response = "7BE26B2DAA223DDBDA"
QueryTable(146).Query = "DFDE03DB6BA4B10F51"
QueryTable(146).Response = "127B7506AC5AECB9F8"
QueryTable(147).Query = "A9BCE33A02B239CCCE"
QueryTable(147).Response = "BD2AFB97A9834F2D1A"
QueryTable(148).Query = "0A40CEC665A38D1613"
QueryTable(148).Response = "463C8265154AC8D67B"
QueryTable(149).Query = "007E6BD99B244EF254"
QueryTable(149).Response = "222156A9FCBF44F413"
QueryTable(150).Query = "34B053D6F8C63312F1"
QueryTable(150).Response = "9DA9B7C91DF1E6D5C6"
QueryTable(151).Query = "17CBEBC0D25B334666"
QueryTable(151).Response = "CEC1F57CF010BEA071"
QueryTable(152).Query = "6DD70B4D5D56DC2819"
QueryTable(152).Response = "71438C311F629B05BC"
QueryTable(153).Query = "78B811263CEE2DE412"
QueryTable(153).Response = "670A4D5E14372872C8"
QueryTable(154).Query = "7F80124AC90694CD67"
QueryTable(154).Response = "31A3927B29F4DA5AB5"
QueryTable(155).Query = "FF740B538A28344D75"
QueryTable(155).Response = "9095C3C77CEEFC2286"
QueryTable(156).Query = "13AB81544DB764D0A7"
QueryTable(156).Response = "772BA5ADB40174FBCB"
QueryTable(157).Query = "7CDF42EA2BAF806075"
QueryTable(157).Response = "18D20FA5247692E3D4"
QueryTable(158).Query = "0BF7964588FB5FAE3B"
QueryTable(158).Response = "A7D8A03929961F0843"
QueryTable(159).Query = "5A1C221D373BE6B316"
QueryTable(159).Response = "21F1985B87ED8E2D64"
QueryTable(160).Query = "FEC2EE337EFD4F18EF"
QueryTable(160).Response = "A4363E3893866588A7"
QueryTable(161).Query = "88506E257A001457BCED"
QueryTable(161).Response = "F041ABA3036F4473CCB5"
QueryTable(162).Query = "6E24064600D30A413BDF"
QueryTable(162).Response = "E416FA42973ED841D7B6"
QueryTable(163).Query = "EEBB96B3001EE1B2A688"
QueryTable(163).Response = "1F08A140C7D01F692D67"
QueryTable(164).Query = "D67A932AF26A41ED8261"
QueryTable(164).Response = "E130762AAA138CF63BFB"
QueryTable(165).Query = "9B5C01C9F9BEA1D55282"
QueryTable(165).Response = "FA7F0E03519FB28CA098"
QueryTable(166).Query = "2806ADE1FDF4D9BE834B"
QueryTable(166).Response = "F4F550949107BEDE7CE1"
QueryTable(167).Query = "5ED1A514B304B236B21D"
QueryTable(167).Response = "400A920B0B89FBFD185D"
QueryTable(168).Query = "6C3C489E5B5871E968A8"
QueryTable(168).Response = "4DED5ECB4ADCBBD5E503"
QueryTable(169).Query = "64A8A6B90A6C734AB594"
QueryTable(169).Response = "0C2E4B15077C47105ADB"
QueryTable(170).Query = "C0520F71AD4350A75991"
QueryTable(170).Response = "9DA61CBF4F5A9479A5F2"
QueryTable(171).Query = "B5D7F86E2D7B26D85511"
QueryTable(171).Response = "EEB86461DD2D92F291E7"
QueryTable(172).Query = "7AD6A67B5FFFEDA61F3B"
QueryTable(172).Response = "FF70A805ACB9D11B801E"
QueryTable(173).Query = "FC6A41CEA0B11592D3DC"
QueryTable(173).Response = "1722CDAFE2644BFB554A"
QueryTable(174).Query = "8D8332624E81B1BA0752"
QueryTable(174).Response = "A9708812F9473BCD8800"
QueryTable(175).Query = "8D7C02E782C3F8E03ED5"
QueryTable(175).Response = "CD6AF97BDEEC2B374642"
QueryTable(176).Query = "2A511606ADC7F9D90686"
QueryTable(176).Response = "5B53A7AB936AD99CE38E"
QueryTable(177).Query = "A269F615110797FD3B25"
QueryTable(177).Response = "AB8D97464B0D2F3D3044"
QueryTable(178).Query = "A8CD017226615F471C64"
QueryTable(178).Response = "6302D1DFFFA08908C254"
QueryTable(179).Query = "E642BB146C30AD450231"
QueryTable(179).Response = "895DA5150DF6B3F1246E"
QueryTable(180).Query = "A7A926163E4E3A0EFB68"
QueryTable(180).Response = "A9B78D863EC7C9BB9F35"
QueryTable(181).Query = "3582B151993D0ECEA4D071"
QueryTable(181).Response = "B6DFD3D708824814852065"
QueryTable(182).Query = "94CB35FB0B64B2E951A300"
QueryTable(182).Response = "685450F2EB73326CE37A69"
QueryTable(183).Query = "2182C7AD0A3795BFA14C06"
QueryTable(183).Response = "7BB3B62EE658AB11B4004F"
QueryTable(184).Query = "9ACAE73D24E3B5BB269F77"
QueryTable(184).Response = "FE44CFABB9DA087EC5CD11"
QueryTable(185).Query = "D6F512BE675308F2C4909B"
QueryTable(185).Response = "8A9AD40C7A446EA6D87420"
QueryTable(186).Query = "99EEF405D60B61B2301635"
QueryTable(186).Response = "239C07F102EF130FBDE6BE"
QueryTable(187).Query = "9995D56ECD9F0B114F15EF"
QueryTable(187).Response = "00BF01692D73125EFE70E3"
QueryTable(188).Query = "CC27B5DEDAB3314D273E77"
QueryTable(188).Response = "50D5556775A6BC1900618D"
QueryTable(189).Query = "01B5D435823E559DCD1BC9"
QueryTable(189).Response = "C6423CA273BFFB6F6B11E1"
QueryTable(190).Query = "654964F313FC592977FF00"
QueryTable(190).Response = "D88022150F8ACD8B717B81"
QueryTable(191).Query = "62792F109114260E32A369"
QueryTable(191).Response = "797FB1840DEA2B069404C9"
QueryTable(192).Query = "B6225CE39CB31A9313005F"
QueryTable(192).Response = "CC3F1E743066D36032C9F7"
QueryTable(193).Query = "814E7E99F6C5D4A4ECDAED"
QueryTable(193).Response = "48024C895432004F5846C5"
QueryTable(194).Query = "91C4BD4BA6378D9EA58FC0"
QueryTable(194).Response = "7B0D20031B8D2F5543B728"
QueryTable(195).Query = "955117C4F42FD0E6E68A65"
QueryTable(195).Response = "BF0863D6D7DBBDA2F96D31"
QueryTable(196).Query = "7495B9611EF9A4A0A2ED85"
QueryTable(196).Response = "DAAB392D7E57D9B1E7D4E9"
QueryTable(197).Query = "03A988EC68EBAA06DA9AD5"
QueryTable(197).Response = "64965E421548FAFB5465E5"
QueryTable(198).Query = "1E04FE3F64EC5C25EEDE1D"
QueryTable(198).Response = "6CC8B47867B45D23C58F8B"
QueryTable(199).Query = "61FCA52C3F8001A08DDE56"
QueryTable(199).Response = "C2A9CED8C3567FBD0CAC7F"
QueryTable(200).Query = "9BF1F0A2727F4DB10E8BCB"
QueryTable(200).Response = "DB32839DE416BFE4ED5E28"


End Sub

Public Sub InitQueryTable_MAX100()

QueryTable(1).Query = "5464"
QueryTable(1).Response = "7B3D"
QueryTable(2).Query = "C131"
QueryTable(2).Response = "044B"
QueryTable(3).Query = "BD34"
QueryTable(3).Response = "73D0"
QueryTable(4).Query = "6B1E"
QueryTable(4).Response = "69EB"
QueryTable(5).Query = "6FED"
QueryTable(5).Response = "C55E"
QueryTable(6).Query = "0C14"
QueryTable(6).Response = "A147"
QueryTable(7).Query = "429D"
QueryTable(7).Response = "8A11"
QueryTable(8).Query = "CCAC"
QueryTable(8).Response = "C51D"
QueryTable(9).Query = "35CF"
QueryTable(9).Response = "A4E1"
QueryTable(10).Query = "56F2"
QueryTable(10).Response = "96A5"
QueryTable(11).Query = "F980"
QueryTable(11).Response = "AF91"
QueryTable(12).Query = "73B4"
QueryTable(12).Response = "0671"
QueryTable(13).Query = "572E"
QueryTable(13).Response = "F59B"
QueryTable(14).Query = "97FF"
QueryTable(14).Response = "C006"
QueryTable(15).Query = "9073"
QueryTable(15).Response = "4DE6"
QueryTable(16).Query = "3130"
QueryTable(16).Response = "56CC"
QueryTable(17).Query = "3AFA"
QueryTable(17).Response = "0EE3"
QueryTable(18).Query = "55DF"
QueryTable(18).Response = "E889"
QueryTable(19).Query = "22A5"
QueryTable(19).Response = "28FC"
QueryTable(20).Query = "6AA3"
QueryTable(20).Response = "22EA"
QueryTable(21).Query = "E3C628"
QueryTable(21).Response = "5840B7"
QueryTable(22).Query = "138E1C"
QueryTable(22).Response = "DB7B98"
QueryTable(23).Query = "5FFE5D"
QueryTable(23).Response = "4944E5"
QueryTable(24).Query = "3AA1A6"
QueryTable(24).Response = "656004"
QueryTable(25).Query = "608313"
QueryTable(25).Response = "47AB37"
QueryTable(26).Query = "A88CA4"
QueryTable(26).Response = "7CE1C1"
QueryTable(27).Query = "FD3F04"
QueryTable(27).Response = "7DA7FF"
QueryTable(28).Query = "CA2D15"
QueryTable(28).Response = "6DE5C0"
QueryTable(29).Query = "971196"
QueryTable(29).Response = "A978FB"
QueryTable(30).Query = "59893F"
QueryTable(30).Response = "1820E2"
QueryTable(31).Query = "3D1296"
QueryTable(31).Response = "FDE809"
QueryTable(32).Query = "B256B3"
QueryTable(32).Response = "C559DD"
QueryTable(33).Query = "3497A9"
QueryTable(33).Response = "7FD8F9"
QueryTable(34).Query = "EF195D"
QueryTable(34).Response = "DE948D"
QueryTable(35).Query = "3BCEFD"
QueryTable(35).Response = "9A50BD"
QueryTable(36).Query = "692B01"
QueryTable(36).Response = "EF0BBA"
QueryTable(37).Query = "A10CDA"
QueryTable(37).Response = "7C0A0B"
QueryTable(38).Query = "7310B2"
QueryTable(38).Response = "B1AD7F"
QueryTable(39).Query = "4E2024"
QueryTable(39).Response = "23A301"
QueryTable(40).Query = "5CD92A"
QueryTable(40).Response = "2CB009"
QueryTable(41).Query = "558C31D2"
QueryTable(41).Response = "F73D19E3"
QueryTable(42).Query = "EE93EA7F"
QueryTable(42).Response = "08BA5EBB"
QueryTable(43).Query = "A08E0354"
QueryTable(43).Response = "85AEFFA1"
QueryTable(44).Query = "010FD80E"
QueryTable(44).Response = "629110EB"
QueryTable(45).Query = "676099EE"
QueryTable(45).Response = "62117C6D"
QueryTable(46).Query = "A06A9EBC"
QueryTable(46).Response = "11D41F6A"
QueryTable(47).Query = "E3D5FA0C"
QueryTable(47).Response = "B58D4C42"
QueryTable(48).Query = "077152C0"
QueryTable(48).Response = "527D243A"
QueryTable(49).Query = "F3D5F0CE"
QueryTable(49).Response = "8FEB3635"
QueryTable(50).Query = "56461344"
QueryTable(50).Response = "8A2DAFF7"
QueryTable(51).Query = "2B775C03"
QueryTable(51).Response = "049EF22A"
QueryTable(52).Query = "F790D96E"
QueryTable(52).Response = "9D16C5C2"
QueryTable(53).Query = "54E1EB6F"
QueryTable(53).Response = "0832CDC5"
QueryTable(54).Query = "EF7637A6"
QueryTable(54).Response = "924773D1"
QueryTable(55).Query = "380F51E6"
QueryTable(55).Response = "EF8460A7"
QueryTable(56).Query = "2E83E688"
QueryTable(56).Response = "59A49B14"
QueryTable(57).Query = "6ECA3A4E"
QueryTable(57).Response = "994BA161"
QueryTable(58).Query = "162282C9"
QueryTable(58).Response = "1FBBFD55"
QueryTable(59).Query = "EE03E2DE"
QueryTable(59).Response = "F1DE446C"
QueryTable(60).Query = "522CB2D3"
QueryTable(60).Response = "2A7AF551"
QueryTable(61).Query = "792020F3F4"
QueryTable(61).Response = "A451785905"
QueryTable(62).Query = "B2A4FF4A7F"
QueryTable(62).Response = "DA26378C66"
QueryTable(63).Query = "0D9074AEBA"
QueryTable(63).Response = "F96E0428B8"
QueryTable(64).Query = "18686DF356"
QueryTable(64).Response = "21D1752AB2"
QueryTable(65).Query = "3D8527858F"
QueryTable(65).Response = "04277F4FD3"
QueryTable(66).Query = "3177D32530"
QueryTable(66).Response = "1DD287DACD"
QueryTable(67).Query = "7914D11935"
QueryTable(67).Response = "90DC299886"
QueryTable(68).Query = "B898EF74FA"
QueryTable(68).Response = "0EFD3839E3"
QueryTable(69).Query = "13F53B5B15"
QueryTable(69).Response = "4A4FB5B213"
QueryTable(70).Query = "C3924F8F9A"
QueryTable(70).Response = "8A9C71EA79"
QueryTable(71).Query = "8936B0F57D"
QueryTable(71).Response = "5AD51FE199"
QueryTable(72).Query = "8D50609866"
QueryTable(72).Response = "0F90F698AF"
QueryTable(73).Query = "8E8D9715B8"
QueryTable(73).Response = "1AD8205E19"
QueryTable(74).Query = "A5D02D6CDC"
QueryTable(74).Response = "FAA7D3BCCA"
QueryTable(75).Query = "73ABFD4810"
QueryTable(75).Response = "C02A4C87BD"
QueryTable(76).Query = "7018F359F5"
QueryTable(76).Response = "8E6708754A"
QueryTable(77).Query = "A82F2A3AF6"
QueryTable(77).Response = "48EB0D34EB"
QueryTable(78).Query = "29C6FF15C1"
QueryTable(78).Response = "FC1065AF22"
QueryTable(79).Query = "64377D99EA"
QueryTable(79).Response = "C18016FBF1"
QueryTable(80).Query = "BAF66CF83A"
QueryTable(80).Response = "0BEE96EA49"
QueryTable(81).Query = "BED58CE512CF"
QueryTable(81).Response = "300190111FED"
QueryTable(82).Query = "F5DEC22AEB2D"
QueryTable(82).Response = "510696D558C3"
QueryTable(83).Query = "89FF2E5C52DE"
QueryTable(83).Response = "D4F41C747F79"
QueryTable(84).Query = "2018AB70D6EF"
QueryTable(84).Response = "0E7445C7973C"
QueryTable(85).Query = "765B6F58E166"
QueryTable(85).Response = "4CD025E5A288"
QueryTable(86).Query = "A4F499DEA113"
QueryTable(86).Response = "DA105F7C3ADC"
QueryTable(87).Query = "3B38137A0CF7"
QueryTable(87).Response = "734647E8F959"
QueryTable(88).Query = "8364E9D1E2C6"
QueryTable(88).Response = "3695E608B486"
QueryTable(89).Query = "810D65D7C12F"
QueryTable(89).Response = "A68F9EFECF4D"
QueryTable(90).Query = "93D090D79CF0"
QueryTable(90).Response = "3FEAE082A319"
QueryTable(91).Query = "BDE7ECC435E4"
QueryTable(91).Response = "FD52E183A0FE"
QueryTable(92).Query = "A57A9F788250"
QueryTable(92).Response = "FA2FD2AAE651"
QueryTable(93).Query = "146EE0BA9207"
QueryTable(93).Response = "6356AD930C28"
QueryTable(94).Query = "1E241C47FEE2"
QueryTable(94).Response = "E1874B3C4A2B"
QueryTable(95).Query = "102A61D906F0"
QueryTable(95).Response = "B3D1E5010DBF"
QueryTable(96).Query = "527CFBB846FB"
QueryTable(96).Response = "7282DA3FEBA2"
QueryTable(97).Query = "AF50A2B64B4C"
QueryTable(97).Response = "9ABBA380AC05"
QueryTable(98).Query = "7EF84755A9E2"
QueryTable(98).Response = "8D9A8E58C012"
QueryTable(99).Query = "8BC4EDA5D470"
QueryTable(99).Response = "71DD3C351E14"
QueryTable(100).Query = "94F55AC32ABA"
QueryTable(100).Response = "0145EC27C9F2"
QueryTable(101).Query = "2D5B22CDFA5455"
QueryTable(101).Response = "07A18C092C9408"
QueryTable(102).Query = "17B4A04606DB8E"
QueryTable(102).Response = "A73FE9AE4839B1"
QueryTable(103).Query = "931316F84477CA"
QueryTable(103).Response = "0AF991C4128455"
QueryTable(104).Query = "80FE90EB1A7985"
QueryTable(104).Response = "C3F4F090D3C441"
QueryTable(105).Query = "12CB342DD0B068"
QueryTable(105).Response = "8D45A524079C77"
QueryTable(106).Query = "2D43413A61E6D7"
QueryTable(106).Response = "EBF46C0387E68C"
QueryTable(107).Query = "347EAD6CF6F552"
QueryTable(107).Response = "CA107F65D0273B"
QueryTable(108).Query = "A2291AE907B585"
QueryTable(108).Response = "FBAE77073DAB5C"
QueryTable(109).Query = "BB9834B85184A6"
QueryTable(109).Response = "69FC91976E7027"
QueryTable(110).Query = "03CF3266A58484"
QueryTable(110).Response = "C1E58B09DCA956"
QueryTable(111).Query = "5D34BE5E753FE3"
QueryTable(111).Response = "DC54A390FF9BCA"
QueryTable(112).Query = "625031BDFE3B78"
QueryTable(112).Response = "1A18C15792BA25"
QueryTable(113).Query = "16A95DE7856934"
QueryTable(113).Response = "4454455396A0A7"
QueryTable(114).Query = "5987994398D1E9"
QueryTable(114).Response = "C6D0EE317E1563"
QueryTable(115).Query = "047DE5A62458E4"
QueryTable(115).Response = "F84FE6DBC306C0"
QueryTable(116).Query = "3CEF91A3F90A90"
QueryTable(116).Response = "8F66438261778B"
QueryTable(117).Query = "1360DB77ED5E7E"
QueryTable(117).Response = "B24041EF35355D"
QueryTable(118).Query = "F09663B0DB9272"
QueryTable(118).Response = "869DFE343F8DE5"
QueryTable(119).Query = "B37FDE2BFC8F2E"
QueryTable(119).Response = "C672BA011F958F"
QueryTable(120).Query = "F994FF798CBE15"
QueryTable(120).Response = "480CB50628E2DF"
QueryTable(121).Query = "B0CB68B2A1E1A7B4"
QueryTable(121).Response = "A3CCBFC32B978428"
QueryTable(122).Query = "A4FD2AE80EEB1A27"
QueryTable(122).Response = "4508BC93016BC369"
QueryTable(123).Query = "65E9513276660F97"
QueryTable(123).Response = "1C7699235EF5D0D5"
QueryTable(124).Query = "A2EFD979F51ACB8A"
QueryTable(124).Response = "2F4930B6388CC719"
QueryTable(125).Query = "4CBB3D78F1B3DA51"
QueryTable(125).Response = "47A407CC621721BE"
QueryTable(126).Query = "938D511CE7B4B3B1"
QueryTable(126).Response = "2342F78DFF101AAA"
QueryTable(127).Query = "34E31CCE15DE9666"
QueryTable(127).Response = "5B37EC3DAD7A82F6"
QueryTable(128).Query = "7B603A3CF15499B1"
QueryTable(128).Response = "724928944B928B6A"
QueryTable(129).Query = "24475ADAA21E24D6"
QueryTable(129).Response = "0536A1A19EBC9CB3"
QueryTable(130).Query = "E204BE5F85A7749E"
QueryTable(130).Response = "1B0588D381F42351"
QueryTable(131).Query = "DBA4BD44A93F1CD6"
QueryTable(131).Response = "F6E6B644C63AE624"
QueryTable(132).Query = "29584448539B7FCE"
QueryTable(132).Response = "1738ABE876FEC530"
QueryTable(133).Query = "5BF650EF7B5057DB"
QueryTable(133).Response = "CA884550F49717C5"
QueryTable(134).Query = "F27876FE4D5B32D5"
QueryTable(134).Response = "FD156FBA258E21F2"
QueryTable(135).Query = "E57B5C01A99CF09A"
QueryTable(135).Response = "9661BEB3AE17CEDF"
QueryTable(136).Query = "1DC03DC5A3564689"
QueryTable(136).Response = "FDF95949CC601CA1"
QueryTable(137).Query = "FAAF69DE04B13E07"
QueryTable(137).Response = "30E74A8F502673CE"
QueryTable(138).Query = "CCCFC222C739B5FE"
QueryTable(138).Response = "08A49ACB060AC6C8"
QueryTable(139).Query = "5A4F402C9D5CDB59"
QueryTable(139).Response = "D7F80C342A0C4DB9"
QueryTable(140).Query = "5F816EDB6AEFB689"
QueryTable(140).Response = "0361C2F89C80711F"
QueryTable(141).Query = "E492EBA74A4D2EC09D"
QueryTable(141).Response = "CCD0C92BB8A3D5F12D"
QueryTable(142).Query = "BCEBA9BC101CC8857D"
QueryTable(142).Response = "543DD87365DC530A24"
QueryTable(143).Query = "BC7CEEF57C015136B8"
QueryTable(143).Response = "84C6D414E7AD7ADC97"
QueryTable(144).Query = "940806A9D73ABDDC76"
QueryTable(144).Response = "EE8C5C59AE718FB763"
QueryTable(145).Query = "F405E30B7CDE4A3FE2"
QueryTable(145).Response = "40A96D3956274F884A"
QueryTable(146).Query = "E5CD32D7D5C71EC24E"
QueryTable(146).Response = "0E88CAEB88EB39DD46"
QueryTable(147).Query = "EE3711D53ED5CD1D33"
QueryTable(147).Response = "90EAC7D940CCBE1A45"
QueryTable(148).Query = "D472843B3BFDE16B1D"
QueryTable(148).Response = "9E40C0A60E30B412FC"
QueryTable(149).Query = "467D71BD905FCAF55F"
QueryTable(149).Response = "76F1EFAA0CBADCC58E"
QueryTable(150).Query = "CD35D98A0F67979076"
QueryTable(150).Response = "E25D47FE497AD63AF4"
QueryTable(151).Query = "A22F9DBE0F33EFF03F"
QueryTable(151).Response = "1ECC3490B1A7A6243F"
QueryTable(152).Query = "8D692EBEA4A1486BF0"
QueryTable(152).Response = "A3440373776B11C02F"
QueryTable(153).Query = "087CD9C46CECF66BC1"
QueryTable(153).Response = "D89E0B14A3719D72B7"
QueryTable(154).Query = "4FEA7720E6425610B1"
QueryTable(154).Response = "54DEE55D1176A60331"
QueryTable(155).Query = "4E2605E4C7CA2695B5"
QueryTable(155).Response = "EC2DCEC37584BBF3F7"
QueryTable(156).Query = "3E53B79027674BA6AF"
QueryTable(156).Response = "E478BA13725DE211C0"
QueryTable(157).Query = "BAD5BC8FEB0DF86DDC"
QueryTable(157).Response = "85217DFEC55D30F2BC"
QueryTable(158).Query = "D8FF69FD765AF3CEFD"
QueryTable(158).Response = "616894F8ACA5CE47A9"
QueryTable(159).Query = "1EDF9F06CCBB9AA676"
QueryTable(159).Response = "0095716165EE213125"
QueryTable(160).Query = "6004A023C7C0D8751E"
QueryTable(160).Response = "5D389908B80E6F622C"
QueryTable(161).Query = "1B54369B3E0FDD5C1677"
QueryTable(161).Response = "43A6D908189A14793ADB"
QueryTable(162).Query = "1996E4D8CF24674FACFE"
QueryTable(162).Response = "4F074790EF7C3EAE4CF0"
QueryTable(163).Query = "FD90D10C95BDAAA0B8F1"
QueryTable(163).Response = "0EE5BAC2DE5F8EA76723"
QueryTable(164).Query = "7356A640837889F67E69"
QueryTable(164).Response = "3D964832273C73299AFD"
QueryTable(165).Query = "4E23CC44B2574AC0A0AA"
QueryTable(165).Response = "8954E1D63F3A98548C16"
QueryTable(166).Query = "3F57BAFF91A7368A3918"
QueryTable(166).Response = "AD7AECB85CE1F6FA1892"
QueryTable(167).Query = "EB90F4BFAB6E3B82D4DD"
QueryTable(167).Response = "C1CD07D197452628CC28"
QueryTable(168).Query = "767D22746DA5C76FFFAB"
QueryTable(168).Response = "A19C8C4E7A8D506C6952"
QueryTable(169).Query = "BEAE16245C01DAE628A2"
QueryTable(169).Response = "B5E0249DE07A67A07CB2"
QueryTable(170).Query = "25B5EC03371F0CCAACA4"
QueryTable(170).Response = "621B57760E1CE5067185"
QueryTable(171).Query = "171AAB1A9C155F503F9F"
QueryTable(171).Response = "71B3F66CE498B21F70CD"
QueryTable(172).Query = "ABF20B6419AC0A1C1492"
QueryTable(172).Response = "8AC1CB9A44FAAA51AD64"
QueryTable(173).Query = "130E58C0E4D2A63D715C"
QueryTable(173).Response = "A7CB8FCE0CABB4558B26"
QueryTable(174).Query = "BC267B368408630D8C8D"
QueryTable(174).Response = "B481CE688EB38EBFFA3B"
QueryTable(175).Query = "416FB65B31D33139EDDA"
QueryTable(175).Response = "748A577264C460191D9B"
QueryTable(176).Query = "A773A1D4B880FA86A8E3"
QueryTable(176).Response = "132E9ECFE81576F18D90"
QueryTable(177).Query = "802B544F27B86B0E2064"
QueryTable(177).Response = "0C69A783A17B340796BA"
QueryTable(178).Query = "DAA4F56888AFD5034411"
QueryTable(178).Response = "ACCB01848FC376400979"
QueryTable(179).Query = "2DC2FC5178FE2F3F57A6"
QueryTable(179).Response = "1EE6C8EA5385EF031A7E"
QueryTable(180).Query = "A4D7E43F5D684328C7F6"
QueryTable(180).Response = "EB4BF3310918D89D5ABA"
QueryTable(181).Query = "7F4CF44330CB0797C981A4"
QueryTable(181).Response = "358CE793CE7ED50C5C2D77"
QueryTable(182).Query = "C95D8591530EC6152F7EAB"
QueryTable(182).Response = "FBEB14413EAF0C0FE96A79"
QueryTable(183).Query = "0D9A86F63C182BEEE3B8CD"
QueryTable(183).Response = "EF60837567FD67E5D17CF1"
QueryTable(184).Query = "BD9199A2173DAD631C92BD"
QueryTable(184).Response = "B100553CD1DE61D9807622"
QueryTable(185).Query = "45808AAE2BE93E021A98FD"
QueryTable(185).Response = "A21233D9041C9BC27504AD"
QueryTable(186).Query = "A2ED4C9BDA4A630BA5D443"
QueryTable(186).Response = "0E5756A6B3DF32C77CA15A"
QueryTable(187).Query = "302BA2F8F188F11E8CCF15"
QueryTable(187).Response = "0EFC0AA93EAD95FE8C38BC"
QueryTable(188).Query = "022C4EB604BD143840FFB1"
QueryTable(188).Response = "028FF1A2FE1F0A8AD4C022"
QueryTable(189).Query = "7B6DD93DC504F19CFCA579"
QueryTable(189).Response = "EE5FED4CE5ABEA745C8DB2"
QueryTable(190).Query = "FF19B84CAED9503906B554"
QueryTable(190).Response = "BDDB95D31EF32BF1743FE9"
QueryTable(191).Query = "9EBD15E0396EC1E93C5198"
QueryTable(191).Response = "16396D0FFDED2DA596FE6D"
QueryTable(192).Query = "209F498419BA25ADAD131F"
QueryTable(192).Response = "C670ACD7B1B3212502EADD"
QueryTable(193).Query = "ED9D75A98577426A0EBE83"
QueryTable(193).Response = "889D54E59A69BA87D77667"
QueryTable(194).Query = "B2455B169BDCDBAC2E9507"
QueryTable(194).Response = "C05633D66CE0589D07799E"
QueryTable(195).Query = "6FEA91C84F21A79308F094"
QueryTable(195).Response = "F00D2A2632688D3682153D"
QueryTable(196).Query = "015B5E4E0306879526F3FB"
QueryTable(196).Response = "6BEC310E406F8E50CC75A6"
QueryTable(197).Query = "CB16B06105611F60627E6F"
QueryTable(197).Response = "5C9BFF7731F8285DEFD39D"
QueryTable(198).Query = "BF9C0E4ACADC6150F8D33C"
QueryTable(198).Response = "C8A50C4BA7B1A47A17D123"
QueryTable(199).Query = "C5CC06DA86F9693042C930"
QueryTable(199).Response = "80AD5DC375F7EA8FF72A93"
QueryTable(200).Query = "76E6BFE2E2597CAFB76ADF"
QueryTable(200).Response = "C0FCC425E12B7355A0E787"
End Sub

    Public Function GetQueryData(Index As Integer) As DATAQUERY

    Dim i As Integer
    Dim Answer As DATAQUERY

    For i = 0 To (Len(QueryTable(Index).Query) / 2 - 1)
           Answer.data(i) = Val("&H" + Mid$(QueryTable(Index).Query, i * 2 + 1, 2))
    Next i

    GetQueryData = Answer

    End Function

    Public Function GetResponseData(Index As Integer) As DATAQUERY

    Dim i As Integer
    Dim Answer As DATAQUERY

    For i = 0 To (Len(QueryTable(Index).Response) / 2 - 1)
           Answer.data(i) = Val("&H" + Mid$(QueryTable(Index).Response, i * 2 + 1, 2))
    Next i

    GetResponseData = Answer

    End Function

'===5/8/02 new dongle testing end ===






 Sub INITIALISE_dongle_stuff()

' === 15/05/02 ===
' Note: DongleResponse( ) is for the OLD dongle ( UK-E4BASBH-B )

    'Passwords
    DonglePassword(0) = "6m\)|e`q,,83Cz;f!I}LS'Z6X'Eq%oq<IsE#UlumHeva>jl(8!@:YAC4z)z=>?UgG"
'    DongleResponse(0) = -1
    DonglePassword(1) = "VyH?(0l?tSM5l9/p+$(\AhluARJNNiyN3ioHe='T3\-FjdPZA]8)'?v4W"
'    DongleResponse(1) = 10230
    DonglePassword(2) = "bM0il@=F):G\{(rsVN=q(|(" + Chr$(34) + "Fq(<([a"
'    DongleResponse(2) = 13443
    DonglePassword(3) = "J./i+wfK#\W>Tx4(\r(]VA&k,t"
'    DongleResponse(3) = -11594
    DonglePassword(4) = "&md(i"
'    DongleResponse(4) = -1108
    DonglePassword(5) = "x>]K)b5mxRr\INv" + Chr$(34) + "#j!'1^f" + Chr$(34) + "'%\KaU/|)_2" + Chr$(34) + "n"
'    DongleResponse(5) = -9051
    DonglePassword(6) = "_opuH|D]Eg=E"
'    DongleResponse(6) = -22043
    DonglePassword(7) = "+WfPa+^Y1\Cjp,[2K2)'6u*g8nXh%6<3C[4C?U<e]WH{@UVHJ0@c-\R\)j5?"
'    DongleResponse(7) = 16309
    DonglePassword(8) = "uXO>zZ&yBmRVY|Xzy`Zmms"
'    DongleResponse(8) = 20140
    DonglePassword(9) = "UXq1=JBo.[t-avR%l:[L[o*hn-e\1Lp06b4-F0-^4rtzjzDK"
'    DongleResponse(9) = 11380
    DonglePassword(10) = "Hy,i#Y#cRC1:/uW08K],<x\%$CZ[#LJ&6V/&(jQ-VCq"
'    DongleResponse(10) = -8466
    DonglePassword(11) = "l?0u_EBM"
'    DongleResponse(11) = -5489
    DonglePassword(12) = "'x2S)g7~%/Q*95OswF[\PDh?F%z@iwPFib0ip&_jHfUIkz-C"
'    DongleResponse(12) = -10955
    DonglePassword(13) = "od/VC6[FEs`H" + Chr$(34) + "82C)P0"
'    DongleResponse(13) = 11823
    DonglePassword(14) = ".x>qDf\|%}I1|yPANUYZB3*)9e~+"
'    DongleResponse(14) = -21686
    DonglePassword(15) = "g*Xb)L^5V,W-+|p|p,"
'    DongleResponse(15) = -21273
    DonglePassword(16) = "Ei0>*J}UA}g<EI}.C-0)GY;DqPwY~xF.eu5ji-i["
'    DongleResponse(16) = 5441
    DonglePassword(17) = "%B2X'8om+=I*AS/v60c3gfd2}_>Slp&Yxb)'3;O`&k{L@<ax<YLdoT-E&H1"
'    DongleResponse(17) = -16733
    DonglePassword(18) = "!\]+ac9xBNq\~S|^,xhVf/.eb8/" + Chr$(34) + "-E^y8Q0&"
'    DongleResponse(18) = 20189
    DonglePassword(19) = ";!C':e@{Kfm4u1OYc" + Chr$(34) + "Dt6Va"
'    DongleResponse(19) = -24320
    DonglePassword(20) = "M$v'vE+v>XRr%tH!J"
'    DongleResponse(20) = -24483
    DonglePassword(21) = ",.W`Au4Dnxt0+KWXX~6,wKe<U<4,V,F`cKD'"
'    DongleResponse(21) = -22306
    DonglePassword(22) = "`QP7C^cth[Y-Vq,0{HG\p12HC8Gg`c-OzV`hcx!e4OCc?62NQ=4t"
'    DongleResponse(22) = -26471
    DonglePassword(23) = "Bs]hs" + Chr$(34) + "F@"
'    DongleResponse(23) = -6113
    DonglePassword(24) = "\*i91TO,l8m,.Lie)R,;|OV.}sB4jkzVD90O;h9Zv,[E6TLjY{k9}Yiz<l/.!,)"
'    DongleResponse(24) = -22255
    DonglePassword(25) = "-\.2;TSo*!`4u.viRKX.SC83U&3WyP9H_sh7J(p3/S|NID_-\\RouyS9W=RI"
'    DongleResponse(25) = -29996
    DonglePassword(26) = "-|Mj_$U'Of(oFW.,H88c2]8BP0V.30Y.-kO8\!%H{euz7P'7H3RQe|<22]"
'    DongleResponse(26) = 5421
    DonglePassword(27) = "" + Chr$(34) + "E.!y`8tqJaZj]B(;bk1&T$EqzSq^.SQ}S"
'    DongleResponse(27) = -2601
    DonglePassword(28) = "SYwp'Dg.0K`." + Chr$(34) + "XKLBrz<6QQDr(SB|o=7"
'    DongleResponse(28) = 942
    DonglePassword(29) = "E#.HY5DY"
'    DongleResponse(29) = -5283
    DonglePassword(30) = "7[*`/USzjOK8&Q6dS]=AAH\X-^C=NHpkLAh>K*[Y43Mq6|"
'    DongleResponse(30) = -17899
    DonglePassword(31) = "p<vt<>(!OGw>:6?k%RmE:!,P(Dva,bmXiyd0Lq6-C!M^:)+AbVWFAM"
'    DongleResponse(31) = 15861
    DonglePassword(32) = "+_J1Qm,&Hg_WQ\^uwP(nm*ekG'h03>Pe?^UQBd.c4'Z"
'    DongleResponse(32) = -19068
    DonglePassword(33) = "XA/;1jS2RlIuH8*2zhE&?uLY"
'    DongleResponse(33) = -11255
    DonglePassword(34) = "Jnr" + Chr$(34) + "xW*,X|-^<,1B|$d$3wh&u^6zYUa=]I?rjk-}$"
'    DongleResponse(34) = -22806
    DonglePassword(35) = "A8f+3@u]i2f0BFWUeno?}SJULu@3"
'    DongleResponse(35) = -17086
    DonglePassword(36) = "}BSq6<$]:6=|1'ABgK,M6g=x,L73Gd=eu3?Vq+IxxY" + Chr$(34) + "5:>!1u1FDHP" + Chr$(34) + "wm"
'    DongleResponse(36) = -24979
    DonglePassword(37) = "PH_,]+O5nDJ@nM+cnGg]`4]FF"
'    DongleResponse(37) = -1411
    DonglePassword(38) = "b8|CHnD&y_iGV')38.h_W2"
'    DongleResponse(38) = 9117
    DonglePassword(39) = "Lv4srfj07B/g_bNu" + Chr$(34) + "TR4%tF,wU_hv,as<#00W}}d" + Chr$(34) + "]Bp5J~*1fq"
'    DongleResponse(39) = 21458
    DonglePassword(40) = "H8tzN^)B\Qr2jC_&d0`*Vs#mJI#" + Chr$(34) + "Mln"
'    DongleResponse(40) = 128
    DonglePassword(41) = "xI2]hnT'"
'    DongleResponse(41) = 27378
    DonglePassword(42) = "TGcp6r'd$NAk>=1"
'    DongleResponse(42) = -20675
    DonglePassword(43) = "K6^k#W\!3x4`B3J(~xNGryg{F)M[7CKR;8(&`9Y$Ak^-P4![]"
'    DongleResponse(43) = 22097
    DonglePassword(44) = "TBY%\r#.Q"
'    DongleResponse(44) = 7357
    DonglePassword(45) = "yC23ah<0&%RH^I&#Wd<::ZxYb=7pc;Q]bJt/MGA#I(zWuOH:ovmF^(_%It+<6"
'    DongleResponse(45) = -29731
    DonglePassword(46) = "O^ji-lGp:S)R"
'    DongleResponse(46) = 23172
    DonglePassword(47) = "vs]iiWv.JX0iY(|hZAOOsOOGW;f"
'    DongleResponse(47) = -613
    DonglePassword(48) = "Z=RFnB"
'    DongleResponse(48) = -38
    DonglePassword(49) = "0K(uhjoR=2wbf6rrR6*j:prfm0C9Of+nY3Y#"
'    DongleResponse(49) = 20181
    DonglePassword(50) = "o4n:`:6WRU|]Xsi8&"
'    DongleResponse(50) = -2710
    DonglePassword(51) = "PQ)Tq3%x0.vKI*#Jdc~Cbkx5Lqt,]qV=a4;.P&.Z5kQv<peBr]&Q<woY}5#+S"
'    DongleResponse(51) = -32425
    DonglePassword(52) = "mczv`G@XKu\Zu8h&Ddv-g}/P^P5fu,Z.4{y`0^PADZ"
'    DongleResponse(52) = 29647
    DonglePassword(53) = "Z}f|5"
'    DongleResponse(53) = -5
    DonglePassword(54) = ">Q2~(I&jciMHI#q.o@Q]/2o>"
'    DongleResponse(54) = -3246
    DonglePassword(55) = "2n`a,s*vAnb&<,]!BIP?wB#,hN$1QYaE\<"
'    DongleResponse(55) = -14345
    DonglePassword(56) = "B@X|=%+h3t6[/X]%)N/9VtW7" + Chr$(34) + "_2s" + Chr$(34) + "'#N$ipH::%)kEu.y/0y~ph9}Dr0y"
'    DongleResponse(56) = 2161
    DonglePassword(57) = ")?1/(H!Y5nX``W,w`Y1`" + Chr$(34) + "-@1gH0sBc8%rG]SLWS{3Y%x$0M_*HXo"
'    DongleResponse(57) = -23426
    DonglePassword(58) = "OE[[0p&)j4V&T6"
'    DongleResponse(58) = -24266
    DonglePassword(59) = "6rw$A0Ha.d+i`|" + Chr$(34) + "<jyIAwspT)=*T0.`pF4Jy1tOX%1@Fyb3`kEfO1}y$E0"
'    DongleResponse(59) = 28317
    DonglePassword(60) = "+eXj)~MCs66&>gH"
'    DongleResponse(60) = 19710
    DonglePassword(61) = "bK}w_8dY7utu/#;nI=?@8'Y}|$T"
'    DongleResponse(61) = 28738
    DonglePassword(62) = "O?C[nc9}v2q>e_~QU#T5^Ed+~[OXE,;dlhw:y:t0(\d-,{bUxHr`zS4s"
'    DongleResponse(62) = -10986
    DonglePassword(63) = "d%K<#mFt,!Y4\?7l\3AUAHmu71ML:2JrD0pm9VaHHPxhCy|h4385@U(.QBtB,%V>"
'    DongleResponse(63) = 21972
    DonglePassword(64) = "h'.EyB$a@Qi|02YDZ*7MQ+=<>6j[5%zNfEGvA3F{F"
'    DongleResponse(64) = 8308
    DonglePassword(65) = "F!EZ{3D3k;cb$Wzet=*\M!" + Chr$(34) + ")jD2|0;X[k|" + Chr$(34) + "(S`Em$<S"
'    DongleResponse(65) = -12141
    DonglePassword(66) = "!I3dqb)/P4MJH>pO~^pa+IL@yqS8*$G`qpwKmAXBc"
'    DongleResponse(66) = -4464
    DonglePassword(67) = "%gM9iQS.&"
'    DongleResponse(67) = 29898
    DonglePassword(68) = "Ps.1rifg1%e"
'    DongleResponse(68) = -28621
    DonglePassword(69) = "9t(5yg##hk[;)dAoLg+<.sJDzKv[Eqv9qWR<DMky="
'    DongleResponse(69) = 17573
    DonglePassword(70) = "+?J9|7GQyD1zidqwy"
'    DongleResponse(70) = -1207
    DonglePassword(71) = "u~BB?y4z;#="
'    DongleResponse(71) = 23105
    DonglePassword(72) = "|PH/`E(N%NJaf[T}76hYt~Lb<" + Chr$(34) + "" + Chr$(34) + "OmyyaLYL&z]%C#_RR#J1{gD8:PE]'Jno"
'    DongleResponse(72) = -18087
    DonglePassword(73) = "-i/&GV-=IU<_~"
'    DongleResponse(73) = -16872
    DonglePassword(74) = ")00v&I%H"
'    DongleResponse(74) = -28521
    DonglePassword(75) = "Hsc=8GGZ?E$>lMJA29^60["
'    DongleResponse(75) = -19295
    DonglePassword(76) = "%q?WX0jEuB3Mq" + Chr$(34) + "qk62f8od=8}{YN8]Vt<x^|4y[=r@Y7b/K" + Chr$(34) + "X$Az-VV"
'    DongleResponse(76) = -30294
    DonglePassword(77) = "YThsM<g\?\;U53.]8P7Io*Z0CaX.vk:IIG2_" + Chr$(34) + "'>:M(7j6*[W<nt^`(4qx|]6~Xb"
'    DongleResponse(77) = -6138
    DonglePassword(78) = "7Wvhm8t*:pNG=&6Gue&X[|9E7CaWNK3u1XB|VsFGqM"
'    DongleResponse(78) = 941
    DonglePassword(79) = "4+(pH5f-#@oa&HB," + Chr$(34) + "" + Chr$(34) + "G`NG|w(u|e"
'    DongleResponse(79) = 3669
    DonglePassword(80) = "GMKz<4GtFEReMaF,[PB;9N$_);" + Chr$(34) + "oOXkZ"
'    DongleResponse(80) = -7261
    DonglePassword(81) = "F}KPNk"
'    DongleResponse(81) = -566
    DonglePassword(82) = "Z5_O`U<"
'    DongleResponse(82) = -864
    DonglePassword(83) = "#@k2rc2eD#%&|eP'!hJKHww@" + Chr$(34) + "gx*ZN;XiCpOP,#zHdGO$Kg,B+4haUcrf"
'    DongleResponse(83) = 14157
    DonglePassword(84) = "g?yO;U#/w~lgM[DxQy=hL(;za/zI&~{%Qy"
'    DongleResponse(84) = -2573
    DonglePassword(85) = "[1'w5IMK[I#G1j2JbR3zp"
'    DongleResponse(85) = 1563
    DonglePassword(86) = "GtVp}ps\gFtsD](hb)ikH"
'    DongleResponse(86) = -28082
    DonglePassword(87) = "|Z" + Chr$(34) + ",=qhQ?L=.~" + Chr$(34) + ",RSQRn@?j@;@KC7N,?rF2|YC(-sU[C"
'    DongleResponse(87) = -30570
    DonglePassword(88) = "Ka1UU&:}DU/ATSz%FDPFUN`[yFy)~/6r({yM2*Esz~9mmcO6K"
'    DongleResponse(88) = -11787
    DonglePassword(89) = "E)OcrFQ#,K3_CCi7}lkKn%}Ck5Th^lSe"
'    DongleResponse(89) = 20480
    DonglePassword(90) = "K[\6z|5mQ"
'    DongleResponse(90) = -12667
    DonglePassword(91) = "zK<MiL*jt9QBzx(.M)mVEs%@%E@7:2mMj.^("
'    DongleResponse(91) = -18178
    DonglePassword(92) = "RrV4?>|%nw}[J=LfkE%c#U3LEqCy)PMk@_T0aMgv[>R~E;+'$" + Chr$(34) + "/Hi`h(V"
'    DongleResponse(92) = -15597
    DonglePassword(93) = "]6xrXo$!.U~$Xp"
'    DongleResponse(93) = -20898
    DonglePassword(94) = "f/DVjG^fK(y>apFuCcFCi?<lHf1v#F-V"
'    DongleResponse(94) = 30351
    DonglePassword(95) = "8S]p<i*74);^BbRsz8" + Chr$(34) + "9VGiKhi*L(#$u89;P]d7tzY@X29Z;lT"
'    DongleResponse(95) = -4566
    DonglePassword(96) = "B9O-[Y04FTn$bU@-kEh@{f</R9D>*_)Z(;m~"
'    DongleResponse(96) = -22035
    DonglePassword(97) = "`+sqbmc_USndG3R{{xCPM>iqw^rMS_T"
'    DongleResponse(97) = 14854
    DonglePassword(98) = "FEkqJ;8Ba]'A\;|J#=\TeC0*" + Chr$(34) + "" + Chr$(34) + "|Wy>y>/j"
'    DongleResponse(98) = 24658
    DonglePassword(99) = "KaNi;?KM`{>pw?2[bxl[A>yYGH&@BB)x^AdM|te"
'    DongleResponse(99) = 31382
    DonglePassword(100) = "h=V<x;B}{" + Chr$(34) + "v@-89D~o/r5G'NP=;M}k9J[T@L" + Chr$(34) + "=XEGp|I}5"
'    DongleResponse(100) = 8712
    DonglePassword(101) = "|E.ucEeQt$OHyoIm?<ZtQ`iaL$B:(K[(&Y4W#Ie5m:kJ&&"
'    DongleResponse(101) = -2743
    DonglePassword(102) = "3!lGQZM-ItsJ)H&"
'    DongleResponse(102) = 15692
    DonglePassword(103) = "/9_zMyqkxu8.Zy0T1.2oOvwhW>}" + Chr$(34) + "a\>Fd8[<p^5@Lm}H6icpQFMl*UU(wE"
'    DongleResponse(103) = 28490
    DonglePassword(104) = "|DqiVTaO," + Chr$(34) + "<#iW5STTa5P$Ib5["
'    DongleResponse(104) = 10754
    DonglePassword(105) = "wM9i*7H:zlJ5?A5" + Chr$(34) + "mO'gwHsr5wNss|@TSDB-/b*!GR/P-rh"
'    DongleResponse(105) = -30201
    DonglePassword(106) = "-wc)1yNK7R^H(!R3x2}*X:b6],N5zZvAOK;r>%^YmcAc<3Mm\<yBKqP)4}4}mM|"
'    DongleResponse(106) = 13259
    DonglePassword(107) = "IOVhDY_MNs=EUKLW3,"
'    DongleResponse(107) = 17321
    DonglePassword(108) = "fCkgZU*NhB5uNLAHGSl5hK@,wah\,'y@PcpD"
'    DongleResponse(108) = -7374
    DonglePassword(109) = "6zBl1uG1fGOHgA~4m4TGl`.zSVC"
'    DongleResponse(109) = 3385
    DonglePassword(110) = "jz++FT;}mAy%hR]L0W{?EcB&cL77y\=XOE*:M^sA4l^ZvNST^i;"
'    DongleResponse(110) = -22787
    DonglePassword(111) = "W[2s>G+Nbwk{hHRFr'ou8VcHZ&ld06*djdawQls636Q"
'    DongleResponse(111) = 23462
    DonglePassword(112) = "MAR%~Ulapw9}9$-" + Chr$(34) + "V4W'HkR56xe" + Chr$(34) + "o?*TTJ:e8vJfu.Kq93/Jj}r5t[BiB'mB&"
'    DongleResponse(112) = -11634
    DonglePassword(113) = "5mt}t|$vD)?OLU#?R&dr-cX}o^lN;j6#QiG;WAh@=g#<IZz3vr<-Q'y_ZrMM<ph2_"
'    DongleResponse(113) = -1
    DonglePassword(114) = "}WL@n{piukQVLW|wE%3nAJBL;>~()9"
'    DongleResponse(114) = -16669
    DonglePassword(115) = "uV1cN_Wa3q<syy<C9vt=Q"
'    DongleResponse(115) = 10281
    DonglePassword(116) = "*cL-;U}X$+-$AN"
'    DongleResponse(116) = 19063
    DonglePassword(117) = "^{M#m&`:m)fskz" + Chr$(34) + "_wG@tw^-"
'    DongleResponse(117) = -27375
    DonglePassword(118) = "a<d1=DvH"
'    DongleResponse(118) = 28776
    DonglePassword(119) = "uq>`D&)NggpwQM"
'    DongleResponse(119) = 19026
    DonglePassword(120) = "v0sjVfq~ilk':T39e1O(@S2vB'&24roMX.#2{kv0:C}np.JE]T&Bb"
'    DongleResponse(120) = -360
    DonglePassword(121) = "5=nl>uT;:4@,F*$+K.Vg{QxIpA@m11aM=g>i1&r#wQVs`W_z[L*eE"
'    DongleResponse(121) = 24563
    DonglePassword(122) = "KcT]6hQV4n"
'    DongleResponse(122) = -20982
    DonglePassword(123) = "n5e=^&Fzyc4P^j[X^7+tf<9boc1fW/.9$yLBh$`"
'    DongleResponse(123) = 5241
    DonglePassword(124) = "#.sTA;FP:F]/ADALJ3mP5HaKwuh|Da/j;;a"
'    DongleResponse(124) = -26706
    DonglePassword(125) = "b'TQm~.l}X9~})MLRyh@" + Chr$(34) + "OEoTy)s9S"
'    DongleResponse(125) = 16965
    DonglePassword(126) = "re%CjnJBRbq<K#lx:SZ\%o-.W=b"
'    DongleResponse(126) = -11236
    DonglePassword(127) = "}M>lMM-,No-)zP(b}&f|W<@Ygw6kr`qN8L" + Chr$(34) + "`v{6+fFL+iJ/jh\"
'    DongleResponse(127) = 1699
    DonglePassword(128) = "E&k#wYm0-Ca@Qmh%lo86Y"
'    DongleResponse(128) = 29311
    DonglePassword(129) = "" + Chr$(34) + "q4ZS~nx;vxA*Vui$CxdD#.rtm|v@&sZ5e0.oQ*]SzG,5umF"
'    DongleResponse(129) = 16658
    DonglePassword(130) = "8|Wjo\SF"
'    DongleResponse(130) = -4377
    DonglePassword(131) = "zj9Y7RTLf" + Chr$(34) + "AZs(wpX" + Chr$(34) + "4uw]W3&Zwa;!k4R-&kx]SMfr,3MB{}"
'    DongleResponse(131) = -23421
    DonglePassword(132) = "b(OS_|[xMnMCF\[NW"
'    DongleResponse(132) = 1441
    DonglePassword(133) = "Q:,d1yE20-^5k3nqWv[nxFmS+1tIHYkFaQ*Zz-6TI/-3_1'Y'B"
'    DongleResponse(133) = -2782
    DonglePassword(134) = "^T7=)N|\d|^K~5zfi7(#;ulxI" + Chr$(34) + "}&?q,z" + Chr$(34) + "p6WOR'k$cI2"
'    DongleResponse(134) = -19186
    DonglePassword(135) = "}r4bIUJDNE6K7HK'qP]G:/>l-cPu#c_6xT^ywC."
'    DongleResponse(135) = 22663
    DonglePassword(136) = "(10lf8vo{jsE3k/$33yp@:Svz97%-;jpXJt!y8+b,8pM$wg6\l:I$uLay<"
'    DongleResponse(136) = 26880
    DonglePassword(137) = "QL(" + Chr$(34) + "<yMW"
'    DongleResponse(137) = -22256
    DonglePassword(138) = "84`^21_44@iGCnZ/RD*XmwIt?TN6#TE)&t<g)o"
'    DongleResponse(138) = 27933
    DonglePassword(139) = ",+lrZ;S0U7N*]NJBb+LrbfzK#h+X#fRUd.+"
'    DongleResponse(139) = 6094
    DonglePassword(140) = "q<4%1U~,tp,%N8Ymf" + Chr$(34) + "{yyk[8O%QxVd%v4<-ISubmAl8\i,"
'    DongleResponse(140) = 5690
    DonglePassword(141) = "{Dfr#DT5v3t*k18C2]4XM1_Dp%;*>)/c" + Chr$(34) + "]mynpI_704S?)-s/8"
'    DongleResponse(141) = -1112
    DonglePassword(142) = "pAZ:ls$0y5|f6VW1MRtphrHs-]niBM#E51F+iT?C*Py=+>4ae;%5<O~o/xwXW"
'    DongleResponse(142) = 5187
    DonglePassword(143) = "!Ky-%X&Rzkl@U6,uT<" + Chr$(34) + "7nv#fx|r@v|?m6&@D),ILzCDp>w&~^rWUi"
'    DongleResponse(143) = -578
    DonglePassword(144) = "f::3X#?BE/d*^&',a" + Chr$(34) + "C{(7hm`"
'    DongleResponse(144) = -1642
    DonglePassword(145) = "SAeFG!6g" + Chr$(34) + "GX.b+RL"
'    DongleResponse(145) = -19115
    DonglePassword(146) = "N4@P{Yn5Q2[I<j<8Zeu;3:n5R.oT^TH\,8"
'    DongleResponse(146) = -22281
    DonglePassword(147) = "x*zWQw}2ph/mlXs8vB\0tJQ<"
'    DongleResponse(147) = -23659
    DonglePassword(148) = "G5`-v'=ttRv?(h&U]UL72KVnt,G)"
'    DongleResponse(148) = -29720
    DonglePassword(149) = "#\oh>"
'    DongleResponse(149) = -142
    DonglePassword(150) = "#hy>ui+N_ulH(@ksX;d%V5N6m>v%@-[a@X2SP/MzU=EGx6\fsk|l_fS:-"
'    DongleResponse(150) = -10494
    DonglePassword(151) = "O^Jvxh"
'    DongleResponse(151) = -287
    DonglePassword(152) = ";$.6GG|%"
'    DongleResponse(152) = -215
    DonglePassword(153) = "i3`gqb.^O$]p"
'    DongleResponse(153) = -22506
    DonglePassword(154) = "2&<$"
'    DongleResponse(154) = -598
    DonglePassword(155) = "bR#gChj:Y|`5/H;wm=9FGfUr${j]Yml0ZC_?/xe*/9"
'    DongleResponse(155) = 21163
    DonglePassword(156) = ">v{/gu;"
'    DongleResponse(156) = -88
    DonglePassword(157) = "Od{KMDPT=nY+Bfp]-Nb.hG]FxCcbpL+__)7L!3p"
'    DongleResponse(157) = -28556
    DonglePassword(158) = "khR5EZCg"
'    DongleResponse(158) = -12419
    DonglePassword(159) = "Gz>JvTGeRT2B" + Chr$(34) + "28TULpZ#H9lHPFHf@g&x[O:3EsCJFtJ+u,l0TvKE"
'    DongleResponse(159) = 1568
    DonglePassword(160) = "h|oKzX_0JO,Y!HjTZN-0,99`')wFwanpH{&90(w8X"
'    DongleResponse(160) = -5510
    DonglePassword(161) = "|l*$4KDg[f,z!wr&sw1MqR}I*p<#%" + Chr$(34) + "}YS-(;~Pzn{&4ar%>D'$x{"
'    DongleResponse(161) = 31174
    DonglePassword(162) = ":#}+"
'    DongleResponse(162) = -9
    DonglePassword(163) = "XUoHWoN3>a+.B1~[[GJ0`OZP?sn%GizMx}/]Evjh;l"
'    DongleResponse(163) = -27330
    DonglePassword(164) = "b!<Nav,Q5]5X.av:(7M=(5rAM`kw%AtWGHAG_dUa8w9"
'    DongleResponse(164) = 21406
    DonglePassword(165) = "-7`i00-%eVgo^+Pl?^M:':Ban28B%1]Bi,bg)oi)e2.$!?(BsD(.J4D6M"
'    DongleResponse(165) = 14703
    DonglePassword(166) = "dwD$Hx?MQZZxHAA3W@;!TL" + Chr$(34) + "_6X?rGxr"
'    DongleResponse(166) = -363
    DonglePassword(167) = "Qk%2XS)N^gS*+dpgX.:jy%|`M$OxQc>..El:5'@<]X]|s,T$"
'    DongleResponse(167) = 30133
    DonglePassword(168) = "qg@=}'VAqd'cK(n/+Lx>H@%" + Chr$(34) + "Qb8gKa2OvQ=]6UNBOkt#)I|"
'    DongleResponse(168) = 29092
    DonglePassword(169) = "AiqvQ"
'    DongleResponse(169) = -5486
    DonglePassword(170) = "v=zOUiZWBJb@A" + Chr$(34) + "]fY%@P)OQ`dH"
'    DongleResponse(170) = 31451
    DonglePassword(171) = "Sa&8i&$=lUzQOc)Z},H0t&zj[=/\)6f$S("
'    DongleResponse(171) = 30549
    DonglePassword(172) = "QhZf{Hu@(cF{(:;w`<" + Chr$(34) + "Y9w4#Bx]V=T`'8]h{B,9y2QXZOzOS*"
'    DongleResponse(172) = -2827
    DonglePassword(173) = "\R04I.WP.IcA0}A"
'    DongleResponse(173) = 16384
    DonglePassword(174) = "]6r`;nh2IU9mF=@{'9t7]R5)qpad1Xb3G'vrZ~b!(Aj]_"
'    DongleResponse(174) = 28346
    DonglePassword(175) = "<kL?rlQ72z{C[JCs"
'    DongleResponse(175) = -22499
    DonglePassword(176) = "I7~vOwOv8`NN5zE"
'    DongleResponse(176) = -2692
    DonglePassword(177) = "" + Chr$(34) + "}9Naa;DFZPx>Ftna9V8@rB=4GPg?ui|rlY/_mbbbzl#)x^H7]=A/*Xj/"
'    DongleResponse(177) = 28149
    DonglePassword(178) = "z" + Chr$(34) + "d2IR`'q:QUR&,0k<V[5uCUeYf"
'    DongleResponse(178) = -2094
    DonglePassword(179) = "99DpC]6KEHbvX9$WjCz"
'    DongleResponse(179) = -17035
    DonglePassword(180) = "!/R,W2A.EHs%C[/r9)XA`U5P>"
'    DongleResponse(180) = -5475
    DonglePassword(181) = "}QPRNf{!g)l1MTl5A(zkg]1dZT@qQ#UK1&"
'    DongleResponse(181) = 12474
    DonglePassword(182) = "J}J4:" + Chr$(34) + "1uh" + Chr$(34) + "yQEx#'aN>\z'aE6sXrLBmM#g[9@*%vr9N%BGRd"
'    DongleResponse(182) = -20412
    DonglePassword(183) = "Y$04dJqUWEg{?"
'    DongleResponse(183) = 5128
    DonglePassword(184) = "j1'I_ZjBrFO?xPU:y5G{>rELkhvUhCpb"
'    DongleResponse(184) = 12245
    DonglePassword(185) = "l2%xQqgk9}[b81si`1;YCnG6g.:H'e-4mGX[K+teB$Rs!"
'    DongleResponse(185) = -9507
    DonglePassword(186) = "M1,V)mw1IV_9MfIE,"
'    DongleResponse(186) = 1865
    DonglePassword(187) = "8m=JbWu4" + Chr$(34) + "0sEEPb[e"
'    DongleResponse(187) = 28799
    DonglePassword(188) = "/_`bKum~^78Z;0A#6&erskZ/SW6a$|"
'    DongleResponse(188) = -10377
    DonglePassword(189) = "pZ=gL*/Y$97n5OOIN^Dr*8M6Aj;kwK@lc{NK]hT1!A"
'    DongleResponse(189) = 16037
    DonglePassword(190) = "/mxF'"
'    DongleResponse(190) = -22
    DonglePassword(191) = "E).3<EE2F/--6niH!rk-f:&=_D8As#gg2mU&;b"
'    DongleResponse(191) = 27306
    DonglePassword(192) = "%/n93^\'EP+mo>KhA%o0aZT%Do62XPUJBm}5%#u6"
'    DongleResponse(192) = 26698
    DonglePassword(193) = "rT_h:oR6+vgH+Ks-}"
'    DongleResponse(193) = 21239
    DonglePassword(194) = "8,GQ4;#UXjif$k" + Chr$(34) + "SCrn9~oTKVqQ^nf-<jS$w(Bd_]K8xZT"
'    DongleResponse(194) = 256
    DonglePassword(195) = "@GG7y0`(bB{i,e/Q;"
'    DongleResponse(195) = 29792
    DonglePassword(196) = "|q~5sA:N&1" + Chr$(34) + "jj<}0%Z~M<SE&Xg"
'    DongleResponse(196) = 18606
    DonglePassword(197) = "[s5UCYH-GPv:Ce)E@/0DwwMne@{Hzr}d:V:&|T6fIsLV"
'    DongleResponse(197) = 13307
    DonglePassword(198) = "^QL#?uIz6CmcLFd?|~_Ksb]Q|n7'_#)I}}]^yS&"
'    DongleResponse(198) = -21855
    DonglePassword(199) = "(mG3@p2K(?w{MT98&g" + Chr$(34) + ")R\d1r&@}_*.!^A|zX"
'    DongleResponse(199) = 31375
    DonglePassword(200) = "cZ<OuW~nkMwSM<+>D)M5yn0ZX"
'    DongleResponse(200) = 6109

 End Sub
 Sub INITIALISE_dongle_2_stuff()
 
 
' === 15/05/02 ===
' Note: Dongle_2_Response( ) is for the NEW dongle  ( RN-4CIIBC-B )

Dongle_2_Response(0) = -1
Dongle_2_Response(1) = 17478
Dongle_2_Response(2) = 30830
Dongle_2_Response(3) = 32555
Dongle_2_Response(4) = -1829
Dongle_2_Response(5) = -20001
Dongle_2_Response(6) = 1503
Dongle_2_Response(7) = 7783
Dongle_2_Response(8) = -21149
Dongle_2_Response(9) = 21815
Dongle_2_Response(10) = -9373
Dongle_2_Response(11) = -5597
Dongle_2_Response(12) = 4206
Dongle_2_Response(13) = 11654
Dongle_2_Response(14) = 11514
Dongle_2_Response(15) = -2656
Dongle_2_Response(16) = -4144
Dongle_2_Response(17) = -21247
Dongle_2_Response(18) = -8880
Dongle_2_Response(19) = -5096
Dongle_2_Response(20) = 255
Dongle_2_Response(21) = 8026
Dongle_2_Response(22) = 1926
Dongle_2_Response(23) = -27635
Dongle_2_Response(24) = -6306
Dongle_2_Response(25) = 2998
Dongle_2_Response(26) = -28877
Dongle_2_Response(27) = 20339
Dongle_2_Response(28) = 6594
Dongle_2_Response(29) = -9797
Dongle_2_Response(30) = 19457
Dongle_2_Response(31) = 3572
Dongle_2_Response(32) = 491
Dongle_2_Response(33) = 28899
Dongle_2_Response(34) = 31526
Dongle_2_Response(35) = 11756
Dongle_2_Response(36) = 225
Dongle_2_Response(37) = -2881
Dongle_2_Response(38) = 29629
Dongle_2_Response(39) = -30235
Dongle_2_Response(40) = 486
Dongle_2_Response(41) = 25479
Dongle_2_Response(42) = -24448
Dongle_2_Response(43) = -1680
Dongle_2_Response(44) = -11502
Dongle_2_Response(45) = 31143
Dongle_2_Response(46) = -13253
Dongle_2_Response(47) = -14425
Dongle_2_Response(48) = -139
Dongle_2_Response(49) = 13198
Dongle_2_Response(50) = -32116
Dongle_2_Response(51) = -26127
Dongle_2_Response(52) = 16155
Dongle_2_Response(53) = -7
Dongle_2_Response(54) = 9344
Dongle_2_Response(55) = -7303
Dongle_2_Response(56) = 4557
Dongle_2_Response(57) = -18640
Dongle_2_Response(58) = -2152
Dongle_2_Response(59) = -8486
Dongle_2_Response(60) = 3680
Dongle_2_Response(61) = -19346
Dongle_2_Response(62) = -17533
Dongle_2_Response(63) = -28887
Dongle_2_Response(64) = -28416
Dongle_2_Response(65) = 28005
Dongle_2_Response(66) = 23110
Dongle_2_Response(67) = 31159
Dongle_2_Response(68) = 6391
Dongle_2_Response(69) = 20091
Dongle_2_Response(70) = -29278
Dongle_2_Response(71) = -8259
Dongle_2_Response(72) = -13017
Dongle_2_Response(73) = -16422
Dongle_2_Response(74) = -22575
Dongle_2_Response(75) = -14997
Dongle_2_Response(76) = -2753
Dongle_2_Response(77) = 490
Dongle_2_Response(78) = -30702
Dongle_2_Response(79) = 2756
Dongle_2_Response(80) = 22530
Dongle_2_Response(81) = -1001
Dongle_2_Response(82) = -558
Dongle_2_Response(83) = -17920
Dongle_2_Response(84) = -22488
Dongle_2_Response(85) = -24545
Dongle_2_Response(86) = -9343
Dongle_2_Response(87) = 29198
Dongle_2_Response(88) = -10519
Dongle_2_Response(89) = 5521
Dongle_2_Response(90) = -7797
Dongle_2_Response(91) = 7991
Dongle_2_Response(92) = 25283
Dongle_2_Response(93) = 2072
Dongle_2_Response(94) = -19392
Dongle_2_Response(95) = -3978
Dongle_2_Response(96) = 2640
Dongle_2_Response(97) = -5626
Dongle_2_Response(98) = -16726
Dongle_2_Response(99) = 21420
Dongle_2_Response(100) = -11067
Dongle_2_Response(101) = -24781
Dongle_2_Response(102) = 29363
Dongle_2_Response(103) = -7914
Dongle_2_Response(104) = 30936
Dongle_2_Response(105) = 8912
Dongle_2_Response(106) = 26550
Dongle_2_Response(107) = 24463
Dongle_2_Response(108) = -2789
Dongle_2_Response(109) = -3140
Dongle_2_Response(110) = 29195
Dongle_2_Response(111) = -18368
Dongle_2_Response(112) = -1808
Dongle_2_Response(113) = -1
Dongle_2_Response(114) = -7465
Dongle_2_Response(115) = 5083
Dongle_2_Response(116) = 15104
Dongle_2_Response(117) = 28456
Dongle_2_Response(118) = 25019
Dongle_2_Response(119) = -10901
Dongle_2_Response(120) = -8648
Dongle_2_Response(121) = -32161
Dongle_2_Response(122) = 13820
Dongle_2_Response(123) = 4313
Dongle_2_Response(124) = 34
Dongle_2_Response(125) = -2507
Dongle_2_Response(126) = 20433
Dongle_2_Response(127) = 5161
Dongle_2_Response(128) = -29122
Dongle_2_Response(129) = -11896
Dongle_2_Response(130) = -5009
Dongle_2_Response(131) = 31067
Dongle_2_Response(132) = -6397
Dongle_2_Response(133) = -20499
Dongle_2_Response(134) = -22415
Dongle_2_Response(135) = -31660
Dongle_2_Response(136) = -6016
Dongle_2_Response(137) = -21228
Dongle_2_Response(138) = -30769
Dongle_2_Response(139) = 32309
Dongle_2_Response(140) = -8818
Dongle_2_Response(141) = -18710
Dongle_2_Response(142) = 14142
Dongle_2_Response(143) = 19638
Dongle_2_Response(144) = -18176
Dongle_2_Response(145) = -12087
Dongle_2_Response(146) = -10048
Dongle_2_Response(147) = -28841
Dongle_2_Response(148) = -12431
Dongle_2_Response(149) = -173
Dongle_2_Response(150) = 13172
Dongle_2_Response(151) = -249
Dongle_2_Response(152) = -781
Dongle_2_Response(153) = -21566
Dongle_2_Response(154) = -632
Dongle_2_Response(155) = -9287
Dongle_2_Response(156) = -100
Dongle_2_Response(157) = -283
Dongle_2_Response(158) = -8517
Dongle_2_Response(159) = -5249
Dongle_2_Response(160) = 20721
Dongle_2_Response(161) = 32001
Dongle_2_Response(162) = -12
Dongle_2_Response(163) = 2128
Dongle_2_Response(164) = 30432
Dongle_2_Response(165) = 30967
Dongle_2_Response(166) = 30
Dongle_2_Response(167) = -2747
Dongle_2_Response(168) = -13053
Dongle_2_Response(169) = -5797
Dongle_2_Response(170) = -21078
Dongle_2_Response(171) = -16631
Dongle_2_Response(172) = -12380
Dongle_2_Response(173) = -8890
Dongle_2_Response(174) = 29471
Dongle_2_Response(175) = -8313
Dongle_2_Response(176) = -13444
Dongle_2_Response(177) = -18725
Dongle_2_Response(178) = 3039
Dongle_2_Response(179) = 30443
Dongle_2_Response(180) = -5208
Dongle_2_Response(181) = 28684
Dongle_2_Response(182) = 31552
Dongle_2_Response(183) = -4649
Dongle_2_Response(184) = 511
Dongle_2_Response(185) = -7197
Dongle_2_Response(186) = 14851
Dongle_2_Response(187) = 26840
Dongle_2_Response(188) = -25362
Dongle_2_Response(189) = -9385
Dongle_2_Response(190) = -106
Dongle_2_Response(191) = -29824
Dongle_2_Response(192) = -22817
Dongle_2_Response(193) = -32517
Dongle_2_Response(194) = -16032
Dongle_2_Response(195) = -18753
Dongle_2_Response(196) = -18955
Dongle_2_Response(197) = -32765
Dongle_2_Response(198) = 17903
Dongle_2_Response(199) = 25207
Dongle_2_Response(200) = 16615

 End Sub



