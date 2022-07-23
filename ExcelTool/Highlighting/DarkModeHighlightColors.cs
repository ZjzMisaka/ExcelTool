
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Highlighting;
using Microsoft.CodeAnalysis.Classification;
using RoslynPad.Editor;

namespace Highlighting
{
    public class DarkModeHighlightColors : IClassificationHighlightColors
    {
        public HighlightingColor DefaultBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(220, 220, 220)) };

        public HighlightingColor TypeBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(78, 201, 176)) };
        public HighlightingColor InterfaceBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(184, 215, 163)) };
        public HighlightingColor MethodBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(216, 216, 163)) };
        public HighlightingColor CommentBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(87, 166, 74)) };
        public HighlightingColor XmlCommentBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(200, 200, 200)) };
        public HighlightingColor XmlCdataBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(233, 213, 133)) };
        public HighlightingColor XmlCommentCommentBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(96, 139, 78)) };
        public HighlightingColor KeywordBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(86, 156, 214)) };
        public HighlightingColor ControlKeywordBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(220, 220, 220)) };
        public HighlightingColor PreprocessorKeywordBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(155, 155, 155)) };
        public HighlightingColor StringBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(214, 157, 133)) };
        public HighlightingColor StringEscapeBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(255, 214, 143)) };
        public HighlightingColor BraceMatchingBrush { get; protected set; } = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.FromRgb(220, 220, 220)), Background = new SimpleHighlightingBrush(Color.FromRgb(14, 69, 131)) };
        public HighlightingColor StaticSymbolBrush { get; protected set; } = new HighlightingColor { FontWeight = FontWeights.Bold };

        public const string BraceMatchingClassificationTypeName = "brace matching";

        private readonly Lazy<ImmutableDictionary<string, HighlightingColor>> _map;

        public DarkModeHighlightColors()
        {
            _map = new Lazy<ImmutableDictionary<string, HighlightingColor>>(() => new Dictionary<string, HighlightingColor>
            {
                [ClassificationTypeNames.ClassName] = AsFrozen(TypeBrush),
                [ClassificationTypeNames.StructName] = AsFrozen(TypeBrush),
                [ClassificationTypeNames.InterfaceName] = AsFrozen(InterfaceBrush),
                [ClassificationTypeNames.DelegateName] = AsFrozen(TypeBrush),
                [ClassificationTypeNames.EnumName] = AsFrozen(InterfaceBrush),
                [ClassificationTypeNames.ModuleName] = AsFrozen(TypeBrush),
                [ClassificationTypeNames.TypeParameterName] = AsFrozen(InterfaceBrush),
                [ClassificationTypeNames.MethodName] = AsFrozen(MethodBrush),
                [ClassificationTypeNames.Comment] = AsFrozen(CommentBrush),
                [ClassificationTypeNames.StaticSymbol] = AsFrozen(StaticSymbolBrush),
                [ClassificationTypeNames.XmlDocCommentAttributeName] = AsFrozen(XmlCommentBrush),
                [ClassificationTypeNames.XmlDocCommentAttributeQuotes] = AsFrozen(XmlCommentBrush),
                [ClassificationTypeNames.XmlDocCommentAttributeValue] = AsFrozen(XmlCommentBrush),
                [ClassificationTypeNames.XmlDocCommentCDataSection] = AsFrozen(XmlCdataBrush),
                [ClassificationTypeNames.XmlDocCommentComment] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.XmlDocCommentDelimiter] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.XmlDocCommentEntityReference] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.XmlDocCommentName] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.XmlDocCommentProcessingInstruction] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.XmlDocCommentText] = AsFrozen(XmlCommentCommentBrush),
                [ClassificationTypeNames.Keyword] = AsFrozen(KeywordBrush),
                [ClassificationTypeNames.ControlKeyword] = AsFrozen(ControlKeywordBrush),
                [ClassificationTypeNames.PreprocessorKeyword] = AsFrozen(PreprocessorKeywordBrush),
                [ClassificationTypeNames.StringLiteral] = AsFrozen(StringBrush),
                [ClassificationTypeNames.VerbatimStringLiteral] = AsFrozen(StringBrush),
                [ClassificationTypeNames.StringEscapeCharacter] = AsFrozen(StringEscapeBrush),
                [ClassificationTypeNames.StringEscapeCharacter] = AsFrozen(StringEscapeBrush),
                [BraceMatchingClassificationTypeName] = AsFrozen(BraceMatchingBrush)
            }.ToImmutableDictionary());
        }

        protected virtual ImmutableDictionary<string, HighlightingColor> GetOrCreateMap()
        {
            return _map.Value;
        }

        public HighlightingColor GetBrush(string classificationTypeName)
        {
            GetOrCreateMap().TryGetValue(classificationTypeName, out var brush);
            return brush ?? AsFrozen(DefaultBrush);
        }

        private static HighlightingColor AsFrozen(HighlightingColor color)
        {
            if (!color.IsFrozen)
            {
                color.Freeze();
            }
            return color;
        }
    }
}