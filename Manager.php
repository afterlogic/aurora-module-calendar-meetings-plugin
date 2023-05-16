<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper as MeetingsHelper;
use Aurora\System\Api;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2021, Afterlogic Corp.
 */
class Manager extends \Aurora\Modules\Calendar\Manager
{
	/**
	 * Processing response to event invitation. [Aurora only.](http://dev.afterlogic.com/aurora)
	 *
	 * @param string $sUserPublicId
	 * @param string $sCalendarId Calendar ID
	 * @param string $sEventId Event ID
	 * @param string $sAttendee Attendee identified by email address
	 * @param string $sAction Appointment actions. Accepted values:
	 *		- "ACCEPTED"
	 *		- "DECLINED"
	 *		- "TENTATIVE"
	 *
	 * @return bool
	 */
	public function updateAppointment($sUserPublicId, $sCalendarId, $sEventId, $sAttendee, $sAction)
	{
		$oResult = null;

		$aData = $this->oStorage->getEvent($sUserPublicId, $sCalendarId, $sEventId);
		if ($aData !== false)
		{
			$oVCal = $aData['vcal'];
			$oVCal->METHOD = 'REQUEST';
			return $this->appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $oVCal->serialize());
		}

		return $oResult;
	}

	/**
	 * Allows for responding to event invitation (accept / decline / tentative). [Aurora only.](http://dev.afterlogic.com/aurora)
	 *
	 * @param int|string $sUserPublicId Account object
	 * @param string $sAttendee Attendee identified by email address
	 * @param string $sAction Appointment actions. Accepted values:
	 *		- "ACCEPTED"
	 *		- "DECLINED"
	 *		- "TENTATIVE"
	 * @param string $sCalendarId Calendar ID
	 * @param string $sData ICS data of the response
	 * @param bool $bExternal If **true**, it is assumed attendee is external to the system
	 *
	 * @return bool
	 */
	public function appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $sData, $bExternal = false)
	{
		$oUser = null;
		$oAttendeeUser = null;
		$oDefaultUser = null;
		$bDefaultAccountAsEmail = false;
		$bIsDefaultAccount = false;

		$oVEvent = null;
		$oVCal = \Sabre\VObject\Reader::read($sData);

		if ($oVCal) {
			$sMethod = $sMethodOriginal = (string) $oVCal->METHOD;
			$aVEvents = $oVCal->getBaseComponents('VEVENT');

			if (isset($aVEvents) && count($aVEvents) > 0) {
				$oVEvent = $aVEvents[0];

				$systemAttendee = null;
				$InformatikProjectsModule = Api::GetModule('InformatikProjects');
				if ($InformatikProjectsModule) {
					$senderForExternalRecipients = $InformatikProjectsModule->getConfig('SenderForExternalRecipients');
					if (!empty($senderForExternalRecipients)) {
						$oEmail = \MailSo\Mime\Email::Parse($senderForExternalRecipients);
						$systemAttendee = $oEmail->GetEmail();
					}
				}

				if ($systemAttendee) {
					foreach($oVEvent->ATTENDEE as &$oAttendee) {
						$sEmail = str_replace('mailto:', '', strtolower((string)$oAttendee));
						if (strtolower($sEmail) === strtolower($systemAttendee)) {
							$sAttendee = MeetingsHelper::getEmailForInternalUsers($sAttendee);
							break;
						}
					}
				}
			}
		}

		if (isset($sUserPublicId))
		{
			$bDefaultAccountAsEmail = false;
			$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
			$oDefaultUser = $oUser;
		}
		else
		{
			$oAttendeeUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sAttendee);
			if ($oAttendeeUser instanceof \Aurora\Modules\Core\Classes\User)
			{
				$bDefaultAccountAsEmail = false;
				$oDefaultUser = $oAttendeeUser;
			}
			else
			{
				$bDefaultAccountAsEmail = true;
			}
		}
		$oFromAccount = null;
		if ($oDefaultUser && $oDefaultUser->PublicId !== $sAttendee)
		{
			$oMailModule = \Aurora\System\Api::GetModule('Mail');
			if ($oMailModule)
			{
				$aAccounts = $oMailModule->getAccountsManager()->getUserAccounts($oDefaultUser->EntityId);
				if (is_array($aAccounts))
				{
					foreach ($aAccounts as $oAccount)
					{
						if ($oAccount instanceof \Aurora\Modules\Mail\Classes\Account && $oAccount->Email === $sAttendee)
						{
							$oFromAccount = $oAccount;
							break;
						}
					}
				}
			}
		}

		if (!$bDefaultAccountAsEmail && !$bIsDefaultAccount)
		{
			$oCalendar = $this->getDefaultCalendar($oDefaultUser->PublicId);
			if ($oCalendar)
			{
				$sCalendarId = $oCalendar->Id;
			}
		}

		$bResult = false;
		$sEventId = null;

		$sTo = $sSubject = $sBody = $sSummary = '';

		if ($oVEvent) {

			$sEventId = (string)$oVEvent->UID;
			if (isset($oVEvent->SUMMARY))
			{
				$sSummary = (string)$oVEvent->SUMMARY;
			}
			if (isset($oVEvent->ORGANIZER))
			{
				$sTo = str_replace('mailto:', '', strtolower((string)$oVEvent->ORGANIZER));
				$iPos = strpos($sTo, 'principals/');
				if ($iPos !== false)
				{
					$sTo = \trim(substr($sTo, $iPos + 11), '/');
				}
			}
			if (strtoupper($sMethod) === 'REQUEST')
			{
				$sMethod = 'REPLY';
				$sSubject = $sSummary;

				$sPartstat = strtoupper($sAction);
				switch ($sPartstat)
				{
					case 'ACCEPTED':
						$sSubject = 'Accepted: '. $sSubject;
						break;
					case 'DECLINED':
						$sSubject = 'Declined: '. $sSubject;
						break;
					case 'TENTATIVE':
						$sSubject = 'Tentative: '. $sSubject;
						break;
				}

				$sCN = '';
				if (isset($oDefaultUser) && $sAttendee ===  $oDefaultUser->PublicId)
				{
					if (!empty($oDefaultUser->Name))
					{
						$sCN = $oDefaultUser->Name;
					}
					else
					{
						$sCN = $sAttendee;
					}
				}

				// $bFoundAttendee = false;
				foreach($oVEvent->ATTENDEE as &$oAttendee)
				{
					$sEmail = str_replace('mailto:', '', strtolower((string)$oAttendee));
					if (strtolower($sEmail) === strtolower($sAttendee))
					{
						$oAttendee['CN'] = $sCN;
						$oAttendee['PARTSTAT'] = $sPartstat;
						$oAttendee['RESPONDED-AT'] = gmdate("Ymd\THis\Z");
						// $bFoundAttendee = true;
						break;
					}
				}
				// unset($oVEvent->ATTENDEE);

				// if (!$bFoundAttendee) {
				// 	$oVEvent->add('ATTENDEE', 'mailto:'.$sAttendee, array(
				// 		'CN' => $sCN,
				// 		'PARTSTAT' => $sPartstat,
				// 		'RESPONDED-AT' => gmdate("Ymd\THis\Z")
				// 	));
				// }
			}

			$oVCal->METHOD = $sMethod;
			$oVEvent->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));

			$sBody = $oVCal->serialize();

			if ($sCalendarId !== false && $bExternal === false && !$bDefaultAccountAsEmail)
			{
				unset($oVCal->METHOD);
				if (isset($oDefaultUser))
				{
					if (strtoupper($sAction) == 'DECLINED' || strtoupper($sMethod) == 'CANCEL')
					{
						$this->deleteEvent($sAttendee, $sCalendarId, $sEventId);
					} else {
//						if ($oDefaultUser->PublicId !== $sAttendee) {
						$this->oStorage->updateEventRaw($oDefaultUser->PublicId, $sCalendarId, $sEventId, $oVCal->serialize());
					}
				}
			}

			if (strtoupper($sMethodOriginal) == 'REQUEST' && (strtoupper($sAction) !== 'DECLINED'))
			{
				if (empty($sTo))
				{
					throw new \Aurora\Modules\CalendarMeetingsPlugin\Exceptions\Exception(\Aurora\Modules\CalendarMeetingsPlugin\Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer);
				}
				else if (!empty($sBody) && isset($oDefaultUser) && $oDefaultUser instanceof \Aurora\Modules\Core\Classes\User)
				{
					$bResult = \Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper::sendAppointmentMessage($oDefaultUser->PublicId, $sTo, $sSubject, $sBody, $sMethod, '', $oFromAccount, $sAttendee);
				}
				else
				{
					throw new \Aurora\Modules\CalendarMeetingsPlugin\Exceptions\Exception(\Aurora\Modules\CalendarMeetingsPlugin\Enums\ErrorCodes::CannotSendAppointmentMessage);
				}
			}
			else
			{
				$bResult = true;
			}
		}

		if (!$bResult)
		{
			\Aurora\System\Api::Log('Ics Appointment Action FALSE result!', \Aurora\System\Enums\LogLevel::Error);
			if ($sUserPublicId)
			{
				\Aurora\System\Api::Log('Email: ' . $oDefaultUser->PublicId . ', Action: '. $sAction.', Data:', \Aurora\System\Enums\LogLevel::Error);
			}
			\Aurora\System\Api::Log($sData, \Aurora\System\Enums\LogLevel::Error);
		}
		else
		{
			$bResult = $sEventId;
		}

		return $bResult;
	}
}






