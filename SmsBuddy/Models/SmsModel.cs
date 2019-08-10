﻿using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Data;
using LiteDB;
using NullVoidCreations.WpfHelpers.Base;
using NullVoidCreations.WpfHelpers.DataStructures;

namespace SmsBuddy.Models
{
    class SmsModel: NotificationBase, IModel
    {
        string _message;
        ExtendedObservableCollection<string> _mobileNumbers, _mobileNumbersScheduled;
        long _id;
        bool _repeatDaily;
        int _hour, _minute;
        TemplateModel _template;
        SmsGatewayBase _gateway;
        IEnumerable<Doublet<string, string>> _fields;

        public SmsModel()
        {
            _mobileNumbers = new ExtendedObservableCollection<string>();
            _mobileNumbersScheduled = new ExtendedObservableCollection<string>();
        }

        #region properties

        [BsonId(true)]
        public long Id
        {
            get { return _id; }
            set { Set(nameof(Id), ref _id, value); }
        }

        public ExtendedObservableCollection<string> MobileNumbers
        {
            get { return _mobileNumbers; }
            set { Set(nameof(MobileNumbers), ref _mobileNumbers, value); }
        }

        public ExtendedObservableCollection<string> MobileNumbersScheduled
        {
            get { return _mobileNumbersScheduled; }
            set { Set(nameof(MobileNumbersScheduled), ref _mobileNumbersScheduled, value); }
        }

        public TemplateModel Template
        {
            get { return _template; }
            set
            {
                if(Set(nameof(Template), ref _template, value) && 
                    value != null && 
                    value.Fields != null)
                {
                    var fields = new List<Doublet<string, string>>();
                    foreach (var field in value.Fields)
                    {
                        var messageField = new Doublet<string, string>(field, null);
                        messageField.PropertyChanged += (object sender, PropertyChangedEventArgs e) => Message = GetMessage();
                        fields.Add(messageField);
                    }
                        
                    Fields = fields;
                    Message = GetMessage();
                }
            }
        }

        public SmsGatewayBase Gateway
        {
            get { return _gateway; }
            set { Set(nameof(Gateway), ref _gateway, value); }
        }

        public string Message
        {
            get { return _message; }
            private set { Set(nameof(Message), ref _message, value); }
        }

        public bool RepeatDaily
        {
            get { return _repeatDaily; }
            set { Set(nameof(RepeatDaily), ref _repeatDaily, value); }
        }

        public int Hour
        {
            get { return _hour; }
            set { Set(nameof(Hour), ref _hour, value); }
        }

        public int Minute
        {
            get { return _minute; }
            set { Set(nameof(Minute), ref _minute, value); }
        }

        public IEnumerable<Doublet<string, string>> Fields
        {
            get { return _fields; }
            private set { Set(nameof(Fields), ref _fields, value); }
        }

        #endregion

        string GetMessage()
        {
            if (Template == null || Template.Message == null)
                return string.Empty;

            var messageBuilder = new StringBuilder(Template.Message);
            if (Fields != null)
                foreach (var field in Fields)
                    messageBuilder.Replace(string.Format("<<{0}>>", field.First), field.Second);
            return messageBuilder.ToString();
        }

        public bool Delete()
        {
            var db = Shared.Instance.Database;
            var collection = db.GetCollection<SmsModel>();
            return collection.Delete(Id);
        }
        public bool DeleteAll()
        {
            var db = Shared.Instance.Database;
            var collection = db.GetCollection<SmsModel>();
            IEnumerable<SmsModel> datas = collection.FindAll();
            bool isDelete = true;

            foreach(var data in datas)
            {
                if(data.Id != Id)
                {
                    isDelete = collection.Delete(data.Id);
                    if (!isDelete)
                        return false;
                }
            }

            return true;
        }

        public IEnumerable<IModel> Get()
        {
            var db = Shared.Instance.Database;
            return db.GetCollection<SmsModel>().FindAll();
        }

        public bool Save()
        {
            var db = Shared.Instance.Database;
            var collection = db.GetCollection<SmsModel>();
            var isSaved = collection.Update(this);
            if (isSaved)
                return isSaved;

            Id = collection.Insert(this);
            return true;
        }

        public override bool Equals(object obj)
        {
            var other = obj as SmsModel;
            return other != null && other.Id.Equals(Id);
        }

        public override int GetHashCode()
        {
            return Id.ToString().GetHashCode();
        }
    }
}
