<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin\Classes;

use Aurora\System\Api;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2021, Afterlogic Corp.
 *
 * @internal
 * 
 * @package Calendar
 * @subpackage Classes
 */
class Helper
{
	/**
	 * @param string $sUserPublicId
	 * @param string $sTo
	 * @param string $sSubject
	 * @param string $sBody
	 * @param string $sMethod
	 * @param string $sHtmlBody Default value is empty string.
	 *
	 * @throws \Aurora\System\Exceptions\ApiException
	 *
	 * @return \MailSo\Mime\Message
	 */
	public static function sendAppointmentMessage($sUserPublicId, $sTo, $sSubject, $sBody, $sMethod, $sHtmlBody='', $oAccount = null, $sFromEmail = null)
	{
		$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
		if ($oUser instanceof \Aurora\Modules\Core\Classes\User)
		{
			$oAccount = $oAccount ? $oAccount : \Aurora\System\Api::GetModule('Mail')->getAccountsManager()->getAccountUsedToAuthorize($oUser->PublicId);
			if ($oAccount instanceof \Aurora\Modules\Mail\Classes\Account)
			{
				$oFromAccount = null;
				$InformatikProjectsModule = Api::GetModule('InformatikProjects');
				if ($InformatikProjectsModule) {
					$senderForExternalRecipients = $InformatikProjectsModule->getConfig('SenderForExternalRecipients');
					$oToEmail = \MailSo\Mime\Email::Parse($sTo);
					if (!empty($senderForExternalRecipients) && self::isEmailExternal($oToEmail->GetEmail())) {
						$oEmail = \MailSo\Mime\Email::Parse($senderForExternalRecipients);
						$sFromEmail = $oEmail->GetEmail();
						$oUser = \Aurora\Modules\Core\Module::Decorator()->GetUserByPublicId($sFromEmail);
						if ($oUser) {
							$MailModule = Api::GetModule('Mail');
							if ($MailModule) {
								$oFromAccount = $MailModule->getAccountsManager()->getAccountByEmail($sFromEmail, $oUser->Id);
							}
						}
					}
				}

				$oMessage = self::buildAppointmentMessage($sUserPublicId, $sTo, $sSubject, $sBody, $sMethod, $sHtmlBody, $oAccount, $sFromEmail);

				\Aurora\System\Api::Log('IcsAppointmentActionSendOriginalMailMessage');
				return \Aurora\System\Api::GetModule('Mail')->getMailManager()->sendMessage($oAccount, $oMessage, null, '', '', '', [], $oFromAccount);
			}
		}

		return false;
	}

	/**
	 * @param string $sUserPublicId
	 * @param string $sTo
	 * @param string $sSubject
	 * @param string $sBody
	 * @param string $sMethod Default value is **null**.
	 * @param string $sHtmlBody Default value is empty string.
	 *
	 * @return \MailSo\Mime\Message
	 */
	public static function buildAppointmentMessage($sUserPublicId, $sTo, $sSubject, $sBody, $sMethod = null, $sHtmlBody = '', $oAccount = null, $sFromEmail = null)
	{
		$oMessage = null;
		$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
		if (isset($sFromEmail)) {
			$sFrom = $sFromEmail;
		} else {
			$sFrom = $oAccount ? $oAccount->Email : $oUser->PublicId;
		}
		if ($oUser instanceof \Aurora\Modules\Core\Classes\User && !empty($sTo) && !empty($sBody))
		{
			$oMessage = \MailSo\Mime\Message::NewInstance();
			$oMessage->RegenerateMessageId();
			$oMessage->DoesNotCreateEmptyTextPart();

			$oMailModule = \Aurora\System\Api::GetModule('Mail'); 
			$sXMailer = $oMailModule ? $oMailModule->getConfig('XMailerValue', '') : '';
			if (0 < strlen($sXMailer))
			{
				$oMessage->SetXMailer($sXMailer);
			}

			$oMessage
				->SetFrom(\MailSo\Mime\Email::NewInstance($sFrom))
				->SetSubject($sSubject)
			;

			$oMessage->AddHtml($sHtmlBody);

			$oToEmails = \MailSo\Mime\EmailCollection::NewInstance($sTo);
			if ($oToEmails && $oToEmails->Count())
			{
				$oMessage->SetTo($oToEmails);
			}

			if ($sMethod)
			{
				$oMessage->SetCustomHeader('Method', $sMethod);
			}

			$oMessage->AddAlternative('text/calendar', \MailSo\Base\ResourceRegistry::CreateMemoryResourceFromString($sBody),
					\MailSo\Base\Enumerations\Encoding::_8_BIT, null === $sMethod ? array() : array('method' => $sMethod));
		}

		return $oMessage;
	}

	public static function isEmailExternal($email)
	{
		$informatikProjectsModule = \Aurora\System\Api::GetModule('InformatikProjects');
		if ($informatikProjectsModule) {
			$domain = \MailSo\Base\Utils::GetDomainFromEmail($email);
			$internalDomains = $informatikProjectsModule->getConfig('InternalDomains', []);
			return !in_array($domain, $internalDomains);
		}
		return false;
	}

	public static function getEmailForInternalUsers($email) 
	{
		$InformatikProjectsModule = Api::GetModule('InformatikProjects');
		if ($InformatikProjectsModule) {
			$senderForExternalRecipients = $InformatikProjectsModule->getConfig('SenderForExternalRecipients');
			if (!empty($senderForExternalRecipients) && !self::isEmailExternal($email)) {
				$oEmail = \MailSo\Mime\Email::Parse($senderForExternalRecipients);
				$email = $oEmail->GetEmail();
			}
		}

		return $email;
	}

	public static function getDomainForInvitation($email)
	{
		$domainForInvitation = '';
		$calendarMeetingsPluginModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
		if ($calendarMeetingsPluginModule && self::isEmailExternal($email)) {
			$externalDomainForInvitation = $calendarMeetingsPluginModule->getConfig('ExternalDomainForInvitation', '');
			$domainForInvitation = rtrim($externalDomainForInvitation);
		}
		if (empty($domainForInvitation)) {
			$domainForInvitation = rtrim(\MailSo\Base\Http::SingletonInstance()->GetFullUrl(), '\\/ ');
		}
		return $domainForInvitation;
	}

	/**
	 * @param \Aurora\Modules\Calendar\Classes\Event $oEvent
	 * @param string $sAccountEmail
	 * @param string $sAttendee
	 * @param string $sCalendarName
	 * @param string $sStartDate
	 *
	 * @return string
	 */
	public static function createHtmlFromEvent($sEventCalendarId, $sEventId, $sEventLocation, $sEventDescription, $sAccountEmail, $sAttendee, $sCalendarName, $sStartDate)
	{
		$sHtml = '';
		$aValues = array(
			'attendee' => $sAttendee,
			'organizer' => $sAccountEmail,
			'calendarId' => $sEventCalendarId,
			'eventId' => $sEventId
		);

		$aValues['action'] = 'ACCEPTED';
		$sEncodedValueAccept = \Aurora\System\Api::EncodeKeyValues($aValues);
		$aValues['action'] = 'TENTATIVE';
		$sEncodedValueTentative = \Aurora\System\Api::EncodeKeyValues($aValues);
		$aValues['action'] = 'DECLINED';
		$sEncodedValueDecline = \Aurora\System\Api::EncodeKeyValues($aValues);

		$sHref = self::getDomainForInvitation($sAttendee) . '/?invite=';
		$oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
		if ($oCalendarMeetingsModule instanceof \Aurora\System\Module\AbstractModule)
		{
			$sHtml = file_get_contents($oCalendarMeetingsModule->GetPath().'/templates/CalendarEventInvite.html');
			$sHtml = strtr($sHtml, array(
				'{{INVITE/LOCATION}}'	=> $oCalendarMeetingsModule->i18N('LOCATION'),
				'{{INVITE/WHEN}}'		=> $oCalendarMeetingsModule->I18N('WHEN'),
				'{{INVITE/DESCRIPTION}}'	=> $oCalendarMeetingsModule->i18N('DESCRIPTION'),
				'{{INVITE/INFORMATION}}'	=> $oCalendarMeetingsModule->i18N('INFORMATION', array('Email' => $sAttendee)),
				'{{INVITE/ACCEPT}}'		=> $oCalendarMeetingsModule->i18N('ACCEPT'),
				'{{INVITE/TENTATIVE}}'	=> $oCalendarMeetingsModule->i18N('TENTATIVE'),
				'{{INVITE/DECLINE}}'		=> $oCalendarMeetingsModule->i18N('DECLINE'),
				'{{Calendar}}'			=> $sCalendarName.' '.$sAccountEmail,
				'{{Location}}'			=> $sEventLocation,
				'{{Start}}'				=> $sStartDate,
				'{{Description}}'			=> $sEventDescription,
				'{{HrefAccept}}'			=> $sHref.$sEncodedValueAccept,
				'{{HrefTentative}}'		=> $sHref.$sEncodedValueTentative,
				'{{HrefDecline}}'			=> $sHref.$sEncodedValueDecline
			));
		}

		return $sHtml;
	}

	public static function createCustomHtmlFromEvent($sEventLocation, $sEventDescription, $sAccountEmail, $sCalendarName, $sStartDate)
	{
		$sHtml = '';

		$oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
		if ($oCalendarMeetingsModule instanceof \Aurora\System\Module\AbstractModule)
		{
			$sHtml = file_get_contents($oCalendarMeetingsModule->GetPath().'/templates/CalendarEventInvite.html');
			$sHtml = strtr($sHtml, array(
				'{{INVITE/LOCATION}}'	=> $oCalendarMeetingsModule->i18N('LOCATION'),
				'{{INVITE/WHEN}}'		=> $oCalendarMeetingsModule->I18N('WHEN'),
				'{{INVITE/DESCRIPTION}}'	=> $oCalendarMeetingsModule->i18N('DESCRIPTION'),
				'{{INVITE/INFORMATION}}'	=> $oCalendarMeetingsModule->i18N('INFORMATION', array('Email' => '{{Attendee}}')),
				'{{INVITE/ACCEPT}}'		=> $oCalendarMeetingsModule->i18N('ACCEPT'),
				'{{INVITE/TENTATIVE}}'	=> $oCalendarMeetingsModule->i18N('TENTATIVE'),
				'{{INVITE/DECLINE}}'		=> $oCalendarMeetingsModule->i18N('DECLINE'),
				'{{Calendar}}'			=> $sCalendarName.' '.$sAccountEmail,
				'{{Location}}'			=> $sEventLocation,
				'{{Start}}'				=> $sStartDate,
				'{{Description}}'			=> $sEventDescription
			));
		}

		return $sHtml;
	}

	/**
	 * @param string $sUserPublicId
	 * @param string $sTo
	 * @param string $sSubject
	 * @param string $sHtmlBody
	 *
	 * @throws \Aurora\System\Exceptions\ApiException
	 *
	 * @return \MailSo\Mime\Message
	 */
	public static function sendSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody)
	{
		$oMessage = self::buildSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody);
		$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
		if ($oUser instanceof \Aurora\Modules\Core\Classes\User)
		{
			$oAccount = \Aurora\System\Api::GetModule('Mail')->getAccountsManager()->getAccountUsedToAuthorize($oUser->PublicId);
			if ($oMessage && $oAccount instanceof \Aurora\Modules\Mail\Classes\Account)
			{
				try
				{
					\Aurora\System\Api::Log('IcsAppointmentActionSendSelfMailMessage');
					return \Aurora\System\Api::GetModule('Mail')->getMailManager()->sendMessage($oAccount, $oMessage);
				}
				catch (\Aurora\System\Exceptions\ManagerException $oException)
				{
					$iCode = \Core\Notifications::CanNotSendMessage;
					switch ($oException->getCode())
					{
						case Errs::Mail_InvalidRecipients:
							$iCode = \Core\Notifications::InvalidRecipients;
							break;
					}

					throw new \Aurora\System\Exceptions\ApiException($iCode, $oException);
				}
			}
		}

		return false;
	}

	/**
	 * @param string $sUserPublicId
	 * @param string $sTo
	 * @param string $sSubject
	 * @param string $sHtmlBody
	 *
	 * @return \MailSo\Mime\Message
	 */
	public static function buildSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody)
	{
		$oMessage = null;
		$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
		if ($oUser instanceof \Aurora\Modules\Core\Classes\User && !empty($sTo) && !empty($sHtmlBody))
		{
			$oMessage = \MailSo\Mime\Message::NewInstance();
			$oMessage->RegenerateMessageId();
			$oMessage->DoesNotCreateEmptyTextPart();

			$oMailModule = \Aurora\System\Api::GetModule('Mail');
			$sXMailer = $oMailModule ? $oMailModule->getConfig('XMailerValue', '') : '';
			if (0 < strlen($sXMailer))
			{
				$oMessage->SetXMailer($sXMailer);
			}

			$oMessage
				->SetFrom(\MailSo\Mime\Email::NewInstance($oUser->PublicId))
				->SetSubject($sSubject)
			;

			$oMessage->AddHtml($sHtmlBody);

			$oToEmails = \MailSo\Mime\EmailCollection::NewInstance($sTo);
			if ($oToEmails && $oToEmails->Count())
			{
				$oMessage->SetTo($oToEmails);
			}
		}

		return $oMessage;
	}

	public static function createSelfNotificationSubject($sAction, $sEventName)
	{
		$sResult = self::getSelfNotificationActionName($sAction) . ': ' . $sEventName;

		return $sResult;
	}

	public static function createSelfNotificationHtmlBody($sAction, $aEvent, $sEmail, $sCalendarName, $sStartDate)
	{
		$sHtml = '';
		$sActionName = self::getSelfNotificationActionName($sAction);
		$oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
		if ($oCalendarMeetingsModule instanceof \Aurora\System\Module\AbstractModule)
		{
			$sHtml = file_get_contents($oCalendarMeetingsModule->GetPath().'/templates/CalendarEventSelfNotification.html');
			$sHtml = strtr($sHtml, [
				'{{LOCATION}}'			=> $oCalendarMeetingsModule->i18N('LOCATION'),
				'{{WHEN}}'				=> $oCalendarMeetingsModule->I18N('WHEN'),
				'{{DESCRIPTION}}'		=> $oCalendarMeetingsModule->i18N('DESCRIPTION'),
				'{{INFORMATION}}'		=> $oCalendarMeetingsModule->i18N('INFORMATION', ['Email' => $sEmail]),
				'{{REACTION}}'			=> $oCalendarMeetingsModule->i18N('USER_REACTION'),
				'{{Calendar}}'			=> $sCalendarName.' '.$sEmail,
				'{{Location}}'			=> $aEvent['location'],
				'{{Start}}'				=> $sStartDate,
				'{{Description}}'			=> $aEvent['description'],
				'{{Reaction}}'			=> $sActionName
			]);
		}

		return $sHtml;
	}

	public static function getSelfNotificationActionName($sAction)
	{
		$sResult = '';
		$oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
		switch ($sAction)
		{
			case 'ACCEPTED':
				$sResult = $oCalendarMeetingsModule->i18N('ACCEPT');
				break;
			case 'DECLINED':
				$sResult = $oCalendarMeetingsModule->i18N('DECLINE');
				break;
			case 'TENTATIVE':
				$sResult = $oCalendarMeetingsModule->i18N('TENTATIVE');
				break;
		}

		return $sResult;
	}
}
