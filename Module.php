<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\Calendar\Classes\Helper;
use Aurora\Modules\Calendar\Classes\Ics;
use Aurora\Modules\Calendar\Classes\Parser;
use Aurora\Modules\Calendar\Enums\AttendeeStatus;
use Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper as MeetingsHelper;
use Aurora\System\Api;
use Aurora\System\Enums\UserRole;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2021, Afterlogic Corp.
 *
 * @package Modules
 */
class Module extends \Aurora\System\Module\AbstractModule
{
	protected $oManager = null;
	protected $oCalendarManager = null;
	public $oApiFileCache = null;
//	public $oApiCalendarDecorator = null;
//	public $oApiUsersManager = null;

	public function getManager()
	{
		if ($this->oManager === null)
		{
			$this->oManager = new Manager($this);
		}

		return $this->oManager;
	}

	public function getCalendarManager()
	{
		if ($this->oCalendarManager === null)
		{
			$this->oCalendarManager = \Aurora\Modules\Calendar\Module::getInstance()->getManager();
		}

		return $this->oCalendarManager;
	}

	public function getCacheManager()
	{
		if ($this->oApiFileCache === null)
		{
			$this->oApiFileCache = new \Aurora\System\Managers\Filecache();
		}

		return $this->oApiFileCache;
	}

	public function init()
	{
		$this->AddEntries([
			'invite' => 'EntryInvite'
		]);

		$this->subscribeEvent('Calendar::CreateIcs', array($this, 'onCreateIcs'));
		$this->subscribeEvent('CalendarMeetingsPlugin::CreateIcs', array($this, 'onCreateIcs'));
		$this->subscribeEvent('Calendar::populateVCalendar', array($this, 'onPopulateVCalendar'));
		$this->subscribeEvent('Calendar::DeleteEvent', array($this, 'onDeleteEvent'));
		$this->subscribeEvent('Calendar::UpdateEventAttendees', array($this, 'onUpdateEventAttendees'));
		$this->subscribeEvent('CalendarMeetingsPlugin::processICS::UpdateEvent', array($this, 'onProcessICSUpdateEvent'));
		$this->subscribeEvent('CalendarMeetingsPlugin::processICS::Cancel', array($this, 'onProcessICSCancel'));
		$this->subscribeEvent('Calendar::processICS::UpdateEvent', array($this, 'onProcessICSUpdateEvent'));
		$this->subscribeEvent('Calendar::processICS::Cancel', array($this, 'onProcessICSCancel'));
		$this->subscribeEvent('Calendar::processICS::AddAttendeesToResult', array($this, 'onAddAttendeesToResult'));
		$this->subscribeEvent('CalendarMeetingsPlugin::processICS::AddAttendeesToResult', array($this, 'onAddAttendeesToResult'));
		$this->subscribeEvent('Calendar::parseEvent', array($this, 'onParseEvent'));

		$this->aErrors = [
			Enums\ErrorCodes::CannotSendAppointmentMessage => $this->i18N('ERROR_CANNOT_SEND_APPOINTMENT_MESSAGE'),
			Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer => $this->i18N('ERROR_CANNOT_SEND_APPOINTMENT_MESSAGE_NO_ORGANIZER')
		];
	}

	/**
	 * @param int $UserId
	 * @param string $File
	 * @param string $FromEmail
	 * @return boolean
	 * @throws \Aurora\System\Exceptions\ApiException
	 */
	public function UpdateAttendeeStatus($UserId, $File, $FromEmail)
	{
		\Aurora\System\Api::checkUserRoleIsAtLeast(\Aurora\System\Enums\UserRole::NormalUser);
		$sUserPublicId = \Aurora\System\Api::getUserPublicIdById($UserId);
		$mResult = false;

		if (empty($File) || empty($FromEmail))
		{
			throw new \Aurora\System\Exceptions\ApiException(\Aurora\System\Notifications::InvalidInputParameter);
		}
		$sData = $this->getCacheManager()->get($sUserPublicId, $File, '', \Aurora\Modules\Calendar\Module::GetName());
		if (!empty($sData))
		{
			$mResult = $this->getManager()->processICS($sUserPublicId, $sData, $FromEmail, true);
		}

		return $mResult;
	}

	/**
	 * @param int $UserId
	 * @param string $File
	 * @param string $FromEmail
	 * @return boolean
	 * @throws \Aurora\System\Exceptions\ApiException
	 */
	public function GetIcsData($UserId, $File, $FromEmail)
	{
		\Aurora\System\Api::checkUserRoleIsAtLeast(\Aurora\System\Enums\UserRole::NormalUser);
		$sUserPublicId = \Aurora\System\Api::getUserPublicIdById($UserId);

		if (empty($File) || empty($FromEmail))
		{
			throw new \Aurora\System\Exceptions\ApiException(\Aurora\System\Notifications::InvalidInputParameter);
		}
		$sData = $this->getCacheManager()->get($sUserPublicId, $File, '', \Aurora\Modules\Calendar\Module::GetName());
		if (!empty($sData))
		{
			$mResult = $this->getCalendarManager()->processICS($sUserPublicId, $sData, $FromEmail);
			if (is_array($mResult) && !empty($mResult['Action']) && !empty($mResult['Body']))
			{
				$oIcs = Ics::createInstance();

				$oIcs->Uid = $mResult['UID'];
				$oIcs->Sequence = $mResult['Sequence'];
				$oIcs->File = $File;
				$oIcs->Type = !empty($mResult['Action']) ? $mResult['Action'] : '';;
				$oIcs->Attendee = !empty($mResult['Attendee']) ? $mResult['Attendee'] : '';
				$oIcs->Location = !empty($mResult['Location']) ? $mResult['Location'] : '';
				$oIcs->Description = !empty($mResult['Description']) ? $mResult['Description'] : '';
				$oIcs->Summary = !empty($mResult['Summary']) ? $mResult['Summary'] : '';
				$oIcs->When = !empty($mResult['When']) ? $mResult['When'] : '';
				$oIcs->CalendarId = !empty($mResult['CalendarId']) ? $mResult['CalendarId'] : '';

				$oIcs->StartTS = $mResult['StartTS'];
				$oIcs->EndTS = $mResult['EndTS'];

				$oIcs->AttendeeList = $mResult['AttendeeList'];
				$oIcs->Organizer = $mResult['Organizer'];

				$this->broadcastEvent(
					'CreateIcs',
					$mResult,
					$oIcs
				);
				return $oIcs;
			}
		}
		return false;
	}

	/**
	 *
	 * @param int $UserId
	 * @param string $CalendarId
	 * @param string $EventId
	 * @param string $File
	 * @param string $AppointmentAction
	 * @param string $Attendee
	 * @return array|boolean
	 * @throws \Aurora\System\Exceptions\ApiException
	 */
	public function SetAppointmentAction($UserId, $CalendarId, $EventId, $File, $AppointmentAction, $Attendee, $AllEvents = 2, $RecurrenceId = null, $Reminders = [])
	{
		\Aurora\System\Api::checkUserRoleIsAtLeast(\Aurora\System\Enums\UserRole::NormalUser);
		$sUserPublicId = \Aurora\System\Api::getUserPublicIdById($UserId);
		$mResult = false;

		if (empty($AppointmentAction) || empty($CalendarId))
		{
			throw new \Aurora\System\Exceptions\ApiException(\Aurora\System\Notifications::InvalidInputParameter);
		}
		$sData = '';
		if (!empty($EventId))
		{
			$aEventData =  $this->getManager()->getEvent($sUserPublicId, $CalendarId, $EventId);
			if (isset($aEventData) && isset($aEventData['vcal']) && $aEventData['vcal'] instanceof \Sabre\VObject\Component\VCalendar)
			{
				$oVCal = $aEventData['vcal'];
				$oVCal->METHOD = 'REQUEST';
				$sData = $oVCal->serialize();
			}
		}
		else if (!empty($File))
		{
			$sData = $this->getCacheManager()->get($sUserPublicId, $File, '', \Aurora\Modules\Calendar\Module::GetName());
		}
		if (!empty($sData))
		{
			$mProcessResult = $this->getManager()->appointmentAction($sUserPublicId, $Attendee, $AppointmentAction, $CalendarId, $sData, $AllEvents, $RecurrenceId, false, false, $Reminders);
			if ($mProcessResult)
			{
				$mResult = array(
					'Uid' => $mProcessResult
				);
			}
		}

		return $mResult;
	}

	public function EntryInvite()
	{
		$sResult = '';
		$aInviteValues = \Aurora\System\Api::DecodeKeyValues($this->oHttp->GetQuery('invite'));

		$oModuleManager = \Aurora\System\Api::GetModuleManager();
		$sTheme = $oModuleManager->getModuleConfigValue('CoreWebclient', 'Theme');

		if (isset($aInviteValues['organizer']))
		{
			$sOrganizerPublicId = $aInviteValues['organizer'];
			$oOrganizerUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sOrganizerPublicId);
			if (isset($sOrganizerPublicId, $aInviteValues['attendee'], $aInviteValues['calendarId'], $aInviteValues['eventId'], $aInviteValues['action']))
			{
				$oCalendar = $this->getManager()->getCalendar($sOrganizerPublicId, $aInviteValues['calendarId']);
				if ($oCalendar)
				{
					$oEvent = $this->getManager()->getEvent($sOrganizerPublicId, $aInviteValues['calendarId'], $aInviteValues['eventId']);
					if ($oEvent && is_array($oEvent) && 0 < count ($oEvent) && isset($oEvent[0]))
					{
						if (is_string($sResult))
						{
							$sResult = file_get_contents($this->GetPath().'/templates/CalendarEventInviteExternal.html');

							$dt = new \DateTime();
							$dt->setTimestamp($oEvent[0]['startTS']);
							if (!$oEvent[0]['allDay'])
							{
								$sDefaultTimeZone = new \DateTimeZone($oOrganizerUser->DefaultTimeZone);
								$dt->setTimezone($sDefaultTimeZone);
							}

							$sAction = $aInviteValues['action'];
							$sActionColor = 'green';
							$sActionText = '';
							switch (strtoupper($sAction))
							{
								case 'ACCEPTED':
									$sActionColor = 'green';
									$sActionText = 'Accepted';
									break;
								case 'DECLINED':
									$sActionColor = 'red';
									$sActionText = 'Declined';
									break;
								case 'TENTATIVE':
									$sActionColor = '#A0A0A0';
									$sActionText = 'Tentative';
									break;
							}

							$sDateFormat = 'm/d/Y';
							$sTimeFormat = 'h:i A';
							switch ($oOrganizerUser->DateFormat)
							{
								case \Aurora\System\Enums\DateFormat::DDMMYYYY:
									$sDateFormat = 'd/m/Y';
									break;
								case \Aurora\System\Enums\DateFormat::DD_MONTH_YYYY:
									$sDateFormat = 'd/m/Y';
									break;
								default:
									$sDateFormat = 'm/d/Y';
									break;
							}
							switch ($oOrganizerUser->TimeFormat)
							{
								case \Aurora\System\Enums\TimeFormat::F24:
									$sTimeFormat = 'H:i';
									break;
								case \Aurora\System\Enums\DateFormat::DD_MONTH_YYYY:
									\Aurora\System\Enums\TimeFormat::F12;
									$sTimeFormat = 'h:i A';
									break;
								default:
									$sTimeFormat = 'h:i A';
									break;
							}
							$sDateTime = $dt->format($sDateFormat.' '.$sTimeFormat);

							$mResult = array(
								'{{COLOR}}' => $oCalendar->Color,
								'{{EVENT_NAME}}' => $oEvent[0]['subject'],
								'{{EVENT_BEGIN}}' => ucfirst($this->i18N('EVENT_BEGIN')),
								'{{EVENT_DATE}}' => $sDateTime,
								'{{CALENDAR}}' => ucfirst($this->i18N('CALENDAR')),
								'{{CALENDAR_NAME}}' => $oCalendar->DisplayName,
								'{{EVENT_DESCRIPTION}}' => $oEvent[0]['description'],
								'{{EVENT_ACTION}}' => $sActionText,
								'{{ACTION_COLOR}}' => $sActionColor,
								'{{THEME}}' => $sTheme,
							);

							$sResult = strtr($sResult, $mResult);

							// $sStartDate = $dt->format($oEvent[0]['allDay'] ? 'D, F d, o' : 'D, F d, o, H:i');
							// \Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper::sendSelfNotificationMessage(
							// 	$aInviteValues['attendee'],
							// 	$aInviteValues['attendee'],
							// 	\Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper::createSelfNotificationSubject($aInviteValues['action'], $oEvent[0]['subject']),
							// 	\Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper::createSelfNotificationHtmlBody($aInviteValues['action'], $oEvent[0], $aInviteValues['attendee'], $oCalendar->DisplayName, $sStartDate)
							// );
						}
						else
						{
							\Aurora\System\Api::Log('Empty template.', \Aurora\System\Enums\LogLevel::Error);
						}
					}
					else
					{
						\Aurora\System\Api::Log('Event not found.', \Aurora\System\Enums\LogLevel::Error);
					}
				}
				else
				{
					\Aurora\System\Api::Log('Calendar not found.', \Aurora\System\Enums\LogLevel::Error);
				}
				$sAttendee = $aInviteValues['attendee'];
				if (!empty($sAttendee))
				{
					if (isset($oEvent) && isset($oEvent['vcal']) && $oEvent['vcal'] instanceof \Sabre\VObject\Component\VCalendar)
					{
						// Getting current status of prcessed attendee to chech 
						// if sending notification message to organizer is needed
						$sCurrentAttendeeStatus = '';
						foreach ($oEvent['vcal']->VEvent[0]->ATTENDEE as $oVAttendee) {
							$sAttendeeEmail = str_replace('mailto:', '', strtolower($oVAttendee->getValue()));
							$sAttendeedStatus = isset($oVAttendee['PARTSTAT']) ? (string)$oVAttendee['PARTSTAT'] : '';
							if ($sAttendeeEmail === $sAttendee) {
								$sCurrentAttendeeStatus = $sAttendeedStatus;
							}
						}

						// If status is not chenged skipping the message sending
						if ($sCurrentAttendeeStatus !== $sAction) {
							$oVCal = $oEvent['vcal'];
							$oVCal->METHOD = 'REQUEST';
							$sData = $oVCal->serialize();
							$oAttendeeUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sAttendee);
							if ($oAttendeeUser) {
								$sOrganizerPublicId = null;
							}
							$prevState = Api::skipCheckUserRole(true);
							$oAttendeeUser = CoreModule::Decorator()->GetUserByPublicId($sAttendee);
							$bIsExternalAttendee = false;
							if ($oAttendeeUser) {
								$sOrganizerPublicId = $oAttendeeUser->PublicId;
							} else {
								$bIsExternalAttendee = true;
							}

							$this->getManager()->appointmentAction($sOrganizerPublicId, $sAttendee, $sAction, $aInviteValues['calendarId'], $sData, 2, null, true, $bIsExternalAttendee);
							Api::skipCheckUserRole($prevState);
						}
					}
					// $this->getManager()->updateAppointment($sOrganizerPublicId, $aInviteValues['calendarId'], $aInviteValues['eventId'], $sAttendee, $aInviteValues['action']);
				}
			}
		}

		if (!$sResult) {
			$sResult = file_get_contents($this->GetPath().'/templates/CalendarEventInviteExternalNotFound.html');
			$sResult = strtr($sResult, array('{{THEME}}' => $sTheme,));
		}

		return $sResult;
	}

	/**
	 *
	 * @return array
	 */
	public function GetSettings()
	{
		\Aurora\System\Api::checkUserRoleIsAtLeast(\Aurora\System\Enums\UserRole::Anonymous);

		$aSettings = array(
			'AllowAppointments' => $this->getConfig('AllowAppointments', true)
		);

		return $aSettings;
	}

	public function onCreateIcs($aData, &$oIcs)
	{
		$oIcs->Attendee = isset($aData['Attendee']) ? $aData['Attendee'] : null;
		$oIcs->Type = $aData['Action'];
	}

	public function onPopulateVCalendar(&$aData, &$oVEvent)
	{
		$sUserPublicId = $aData['UserPublicId'];
		$oEvent = $aData['Event'];
		$sHtml = isset($aData['appointmentMailBody']) ? $aData['appointmentMailBody'] : null;
		$sSubjectPrexix = isset($aData['isEventUpdate']) && $aData['isEventUpdate'] ? 
			$this->i18N('APPOINTMENT_SUBJECT_UPDATE') : $this->i18N('APPOINTMENT_SUBJECT_NEW');
		$oVCal = $aData['VCal'];
		$sComponentName = $aData['ComponentName'];
		$sComponentIndex = $aData['ComponentIndex'];
		$oVEvent = $oVCal->$sComponentName[$sComponentIndex];
		$oUser = \Aurora\System\Api::getAuthenticatedUser();

		$aAttendees = [];
		$aAttendeeEmails = [];
		$aObjAttendees = [];
		if (isset($oVEvent->ORGANIZER)) {
			$sOwnerEmail = str_replace('mailto:', '', strtolower((string) $oVEvent->ORGANIZER));
			$iPos = strpos($sOwnerEmail, 'principals/');
			if ($iPos !== false) {
				$sOwnerEmail = \trim(substr($sOwnerEmail, $iPos + 11), '/');
			}
		}
		//update only own attendees
		if (!isset($sOwnerEmail) || isset($sOwnerEmail) && $sOwnerEmail === $sUserPublicId) {
			if (isset($oVEvent->ATTENDEE)) {
				$aAttendeeEmails = [];
				foreach ($oEvent->Attendees as $aItem) {
					$sStatus = '';
					switch ($aItem['status']) {
						case AttendeeStatus::Accepted:
							$sStatus = 'ACCEPTED';
							break;
						case AttendeeStatus::Declined:
							$sStatus = 'DECLINED';
							break;
						case AttendeeStatus::Tentative:
							$sStatus = 'TENTATIVE';
							break;
						case AttendeeStatus::Unknown:
							$sStatus = 'NEEDS-ACTION';
							break;
					}

					$aAttendeeEmails[strtolower($aItem['email'])] = $sStatus;
				}

				$aObjAttendees = $oVEvent->ATTENDEE;
				unset($oVEvent->ATTENDEE);
				foreach($aObjAttendees as $oAttendee) {
					$sAttendee = str_replace('mailto:', '', strtolower((string)$oAttendee));
					$oPartstat = $oAttendee->offsetGet('PARTSTAT');
					if (in_array($sAttendee, array_keys($aAttendeeEmails))) {
						if (isset($oPartstat) && (string) $oPartstat === $aAttendeeEmails[$sAttendee]) {
							$oVEvent->add($oAttendee);
							$aAttendees[] = $sAttendee;
						}
					} elseif (!isset($oPartstat) || (isset($oPartstat) && (string)$oPartstat != 'DECLINED')) {
						$oVCal->METHOD = 'CANCEL';
						$sSubject = (string)$oVEvent->SUMMARY . ': Canceled';

						$oVCalResult = clone $oVCal;
						$oVEventResult = $oVCalResult->$sComponentName[$sComponentIndex];

						$oVEventResult->ORGANIZER = $sUserPublicId;

						$this->prepareVEventDataForExternalUser($sUserPublicId, $sAttendee, $oVEventResult);

						MeetingsHelper::sendAppointmentMessage(
							$sUserPublicId,
							$sAttendee,
							$sSubject,
							$oVCalResult->serialize(),
							(string) $oVCalResult->METHOD
						);

						unset($oVCal->METHOD);
					}
				}
			}

			if (is_array($oEvent->Attendees) && count($oEvent->Attendees) > 0) {
				if (!isset($oVEvent->ORGANIZER)) {
					$oVEvent->ORGANIZER = 'mailto:' . $oUser->PublicId;
				}
				foreach($oEvent->Attendees as $oAttendee) {
					if (!in_array($oAttendee['email'], $aAttendees)) {
						$oVEvent->add(
							'ATTENDEE',
							'mailto:' . $oAttendee['email'],
							array(
								'CN' => !empty($oAttendee['name']) ? $oAttendee['name'] : $oAttendee['email'],
								'RSVP' => 'TRUE'
							)
						);
					}
				}
			} else {
				unset($oVEvent->ORGANIZER);
			}

			if (isset($oVEvent->ATTENDEE)) {
				foreach($oVEvent->ATTENDEE as $oAttendee) {							
					$sAttendee = str_replace('mailto:', '', strtolower((string)$oAttendee));

					if (($sAttendee !==  $oUser->PublicId) &&
						(!isset($oAttendee['PARTSTAT']) || 
						(isset($oAttendee['PARTSTAT']) && (string)$oAttendee['PARTSTAT'] !== 'DECLINED'))) {
						
						$sStartDateFormat = $oVEvent->DTSTART->hasTime() ? 'D, d. F, o, H:i' : 'D, d. F, o';
						$sStartDate = Helper::getStrDate($oVEvent->DTSTART, $oUser->DefaultTimeZone, $sStartDateFormat);
						$sStartDate .= $oVEvent->DTSTART->hasTime() ? ' Uhr' : '';

						$sStartWeek = Helper::getStrDate($oVEvent->DTSTART, $oUser->DefaultTimeZone, 'W');
						$sStartDate .= ' ('.$this->i18N('WEEK_SHORT').$sStartWeek.')';

						$oCalendar =  Api::GetModule('Calendar')->GetCalendar($sUserPublicId, $oEvent->IdCalendar);

						$sEventUid = (string)$oVEvent->UID;
						if (!isset($sHtml)) {
							$sFinalHtml = MeetingsHelper::createHtmlFromEvent(
								$oEvent->IdCalendar,
								$sEventUid,
								$oVEvent->Location,
								$oVEvent->Description,
								$oUser->PublicId,
								$sAttendee,
								$oCalendar->DisplayName,
								$sStartDate
							);
						} else {
							$aValues = array(
								'attendee' => $sAttendee,
								'organizer' => $oUser->PublicId,
								'calendarId' => $oEvent->IdCalendar,
								'eventId' => $sEventUid
							);
							$sHref = MeetingsHelper::getDomainForInvitation($sAttendee) . '/?invite=';
					
							$aValues['action'] = 'ACCEPTED';
							$sEncodedValueAccept = \Aurora\System\Api::EncodeKeyValues($aValues);
							$aValues['action'] = 'TENTATIVE';
							$sEncodedValueTentative = \Aurora\System\Api::EncodeKeyValues($aValues);
							$aValues['action'] = 'DECLINED';
							$sEncodedValueDecline = \Aurora\System\Api::EncodeKeyValues($aValues);
					
							$sFinalHtml = strtr($sHtml, array(
								'{{Attendee}}'		=> $sAttendee,
								'{{HrefAccept}}'	=> $sHref . $sEncodedValueAccept,
								'{{HrefTentative}}'	=> $sHref . $sEncodedValueTentative,
								'{{HrefDecline}}'	=> $sHref . $sEncodedValueDecline
							));
						}

						$oVCalResult = clone $oVCal;
						$oVEventResult = $oVCalResult->$sComponentName[$sComponentIndex];

						$oVEventResult->ORGANIZER = $sUserPublicId;

						$this->prepareVEventDataForExternalUser($sUserPublicId, $sAttendee, $oVEventResult);

						$oVCalResult->METHOD = 'REQUEST';

						MeetingsHelper::sendAppointmentMessage(
							$sUserPublicId,
							$sAttendee, 
							$sSubjectPrexix . (string) $oVEventResult->SUMMARY, 
							$oVCalResult->serialize(), 
							(string) $oVCalResult->METHOD, 
							$sFinalHtml
						);
					}
				}
			}
		}
	}

	public function onDeleteEvent($aData, &$oVCal)
	{
		$sUserPublicId = $aData['sUserPublicId'];

		$iIndex = Helper::getBaseVComponentIndex($oVCal->VEVENT);
		if ($iIndex !== false)
		{
			$oVEvent = $oVCal->VEVENT[$iIndex];

			$sOrganizer = (isset($oVEvent->ORGANIZER)) ?
					str_replace('mailto:', '', strtolower((string)$oVEvent->ORGANIZER)) : null;
			$iPos = strpos($sOrganizer, 'principals/');
			if ($iPos !== false)
			{
				$sOrganizer = \trim(substr($sOrganizer, $iPos + 11), '/');
			}

			if (isset($sOrganizer))
			{
				if ($sOrganizer === $sUserPublicId)
				{
					$oDateTimeNow = new \DateTimeImmutable("now");
					$oDateTimeEvent = $oVEvent->DTSTART->getDateTime();
					$oDateTimeRepeat = Helper::getNextRepeat($oDateTimeNow, $oVEvent);
					$bRrule = isset($oVEvent->RRULE);
					$bEventFore = $oDateTimeEvent ? $oDateTimeEvent > $oDateTimeNow : false;
					$bNextRepeatFore = $oDateTimeRepeat ? $oDateTimeRepeat > $oDateTimeNow : false;

					if (isset($oVEvent->ATTENDEE) && ($bRrule ? $bNextRepeatFore : $bEventFore))
					{
						foreach($oVEvent->ATTENDEE as $oAttendee)
						{
							$sAttendee = str_replace('mailto:', '', strtolower((string)$oAttendee));

							$oVCal->METHOD = 'CANCEL';
							$sSubject = (string)$oVEvent->SUMMARY . ': Canceled';

							$oVCalResult = clone $oVCal;
							$oVEventResult = $oVCalResult->VEVENT[$iIndex];

							$oVEventResult->ORGANIZER = $sUserPublicId;

							$this->prepareVEventDataForExternalUser($sUserPublicId, $sAttendee, $oVEventResult);

							MeetingsHelper::sendAppointmentMessage(
								$sUserPublicId,
								$sAttendee,
								$sSubject,
								$oVCalResult->serialize(),
								(string) $oVCal->METHOD
							);

							unset($oVCal->METHOD);
						}
					}
				}
			}
		}
	}

	public function onUpdateEventAttendees($aData, &$oEvent)
	{
		$attendees = $aData['attendees'];
		$owner = isset($aData['owner']) ? $aData['owner'] : '';
		$UserPublicId = isset($aData['UserPublicId']) ? $aData['UserPublicId'] : '';
		if ($owner && $UserPublicId && $owner === $UserPublicId)
		{
			$oEvent->Attendees = @json_decode($attendees, true);
		}
	}

	public function onProcessICSUpdateEvent(&$aData, &$mResult)
	{
		$oVCalResult = $aData['oVCalResult'];
		$oVEventResult = $aData['oVEventResult'];
		$sUserPublicId = $aData['sUserPublicId'];
		$sCalendarId = $aData['sCalendarId'];
		$sEventId = $aData['sEventId'];
		$sMethod = $aData['sMethod'];
		$sequence = $aData['sequence'];
		$sequenceServer = $aData['sequenceServer'];
		$oVEvent = $aData['oVEvent'];

		if (isset($oVEvent->ATTENDEE) && $sequenceServer >= $sequence) {
			foreach ($oVEvent->ATTENDEE as $oAttendee) {
				if (isset($oVEventResult->ATTENDEE)) {
					foreach ($oVEventResult->ATTENDEE as &$oAttendeeResult) {
						if ($oAttendeeResult->getValue() === $oAttendee->getValue()) {
							if (isset($oAttendee['PARTSTAT'])) {
								$oAttendeeResult['PARTSTAT'] = $oAttendee['PARTSTAT']->getValue();
								$oRespondedAt = $oVEvent->{'LAST-MODIFIED'}->getDateTime();
								$oRespondedAt->setTimezone(new \DateTimeZone('UTC'));
								$oAttendeeResult['RESPONDED-AT'] = gmdate("Ymd\THis\Z", $oRespondedAt->getTimestamp());
							}
							break;
						}
					}
				}
			}
		}
		unset($oVCalResult->METHOD);
		$oVEventResult->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));
		$mResult = $this->getManager()->updateEventRaw($sUserPublicId, $sCalendarId, $sEventId, $oVCalResult->serialize());
		$oVCalResult->METHOD = $sMethod;
	}

	public function onProcessICSCancel(&$aData, &$mResult)
	{
		$sCalendarId = $aData['sCalendarId'];
		$sUserPublicId = $aData['sUserPublicId'];
		$sEventId = $aData['sEventId'];

		if ($this->getManager()->deleteEvent($sUserPublicId, $sCalendarId, $sEventId))
		{
			$mResult = true;
		}
	}

	public function onAddAttendeesToResult(&$aData, &$mResult)
	{
		$oVEventResult = $aData['oVEventResult'];
		$aAccountEmails = $aData['aAccountEmails'];
		$sMethod = $aData['sMethod'];

		if (isset($oVEventResult->ATTENDEE))
		{
			foreach($oVEventResult->ATTENDEE as $oAttendee)
			{
				$sAttendee = str_replace('mailto:', '', strtolower((string)$oAttendee));
				if (in_array($sAttendee, $aAccountEmails) && isset($oAttendee['PARTSTAT']) && $sMethod !== 'SAVE')
				{
					$mResult['Attendee'] = $sAttendee;
					$mResult['Action'] = $sMethod . '-' . $oAttendee['PARTSTAT']->getValue();
				}
			}
		}
	}

	public function onParseEvent(&$aData, &$aEvent)
	{
		$oVComponent = $aData['oVComponent'];
		$sOwnerEmail = $aData['sOwnerEmail'];
		$oUser = $aData['oUser'];

		$bIsAppointment = false;
		$aEvent['attendees'] = array();
		if (isset($oVComponent->ATTENDEE))
		{
			$aEvent['attendees'] = Parser::parseAttendees($oVComponent);

			if (isset($oVComponent->ORGANIZER))
			{
				$sOwnerEmail = str_replace('mailto:', '', strtolower((string)$oVComponent->ORGANIZER));
				$iPos = strpos($sOwnerEmail, 'principals/');
				if ($iPos !== false)
				{
					$sOwnerEmail = \trim(substr($sOwnerEmail, $iPos + 11), '/');
				}
				$aData['sOwnerEmail'] = $sOwnerEmail;
			}
			$bIsAppointment = ($oUser instanceof \Aurora\Modules\Core\Classes\User && $sOwnerEmail !== $oUser->PublicId);
		}

		$aEvent['appointment'] = $bIsAppointment;
		$aEvent['appointmentAccess'] = 0;
	}

	public function GetMailMessageBodyForEvent($Location, $Description, $CalendarDisplayName, $StartDate) 
	{
		Api::checkUserRoleIsAtLeast(UserRole::NormalUser);
		$oUser = Api::getAuthenticatedUser();

		return MeetingsHelper::createCustomHtmlFromEvent($Location, $Description, $oUser->PublicId, $CalendarDisplayName, $StartDate);
	}

	/**
	 * The method updated VEVENT data to prevent disclosure of real addresses of internal users.
	 * It replaces domain names of email of internal users.
	 * Also it replaces the organizer address with the email of central account.
	 */
	private function prepareVEventDataForExternalUser($sUserPublicId, $sAttendee, &$oVEvent)
	{
		// replacing real domain of attendees for external attendees
		if (MeetingsHelper::isEmailExternal($sAttendee)) {
			$oVEventAteendeeResult = $oVEvent->ATTENDEE;
			unset($oVEvent->ATTENDEE);

			if ($oVEventAteendeeResult instanceof \Sabre\VObject\Property\ICalendar\CalAddress) {
				foreach ($oVEventAteendeeResult as $oAttendeeResult) {
					$sAttendeeResult = str_replace('mailto:', '', strtolower((string)$oAttendeeResult));
					if (!MeetingsHelper::isEmailExternal($sAttendeeResult)) {
						$sDomainNameForReplacementInInvitations = $this->getConfig('DomainNameForReplacementInInvitations', '');
						if (!empty($sDomainNameForReplacementInInvitations)) {
							$oAttendeeResult->setValue('mailto:' . \MailSo\Base\Utils::GetAccountNameFromEmail($sAttendeeResult) . '@' . $sDomainNameForReplacementInInvitations);
						}
						if (isset($oAttendeeResult['CN'])) {
							$oAttendeeResult['CN'] = \MailSo\Base\Utils::GetAccountNameFromEmail($oAttendeeResult['CN']);
						}
					}

					$oVEvent->add($oAttendeeResult);
				}
			}

			// setting special organizer's email address for external attendees
			$oVEvent->ORGANIZER = MeetingsHelper::getCorrectedSenderEmail($sUserPublicId, $sAttendee);
		}

		// adding sender's friendly name to organizer
		$sFriendlyName = MeetingsHelper::getMainAccountFriendlyName($sUserPublicId);
		if (!empty($sFriendlyName)) {
			$oVEvent->ORGANIZER['CN'] = $sFriendlyName;
		}
	}
}
